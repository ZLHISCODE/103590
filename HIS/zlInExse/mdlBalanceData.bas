Attribute VB_Name = "mdlBalanceData"
Option Explicit
'*********************************************************************************************************************************************
'����:��ȡ��ؽ�����Ϣ����
'����:
'  01����������
'    0101. zlGetBalanceDataErrFromBalanceID:���ݽ���ID��ȡ��ص��쳣����������Ϣ
'    0102. zlGetBalanceDataFromBalanceID:���ݽ���ID��ȡ��صĽ���������Ϣ
'    0103. zlGetSaveThreeDelSwapBatchSQL:����ԭ������Ϣ(clsBalanceItem)������ȡ���������˿���Ϣ��SQL
'    0104. zlGetBalanceItemSQLFromBalanceItem:���ݽ������(clsBalanceItem)��ȡ��������������SQL
'    0105. zlGetSaveThirdSwapDelSQLFromBalanceItem:���ݽ����((clsBalanceItem))���ȡ���������˿��SQL
'    0106. zlThirdDelSwapIsExsistFromBalanceID:���ݽ���ID�ж������˿���Ƿ����
'    0107. zlGet���㷽ʽ:��ȡ���㷽ʽ(��Ӧ�ó���)
'    0108. zlGetClassMoney����ȡ���շ������ܵļ�¼��
'    0109. zlGetRemainderMoneyToPati����ȡ���������Ϣ�����˶���
'    0110. zlGetDefaultHospitalizedDate�����ݲ���ID��ȡ�ϴ���;����ʱ��
'    0111. zlIsCheck�����ѽ���:��鲡���Ƿ����
'    0112. zlGetThirdMoneyInforRecordFromSwapID:���ݽ���ID,��ȡ��ؽ��׵Ľ�����Ϣ��(��ԭʼ�����˽�δ�˽���)
'    0113. zlCheck�������:��鲡���Ƿ����
'    0114. CheckPatiIsVerfy:���ָ�������Ƿ��Ѿ����
'    0115. zlComparePatiNumsIsDiff:�Ƚ�����סԺ�����Ƿ�һ��
'    0116. zlCopyNewFeeData:����ҵ�����͵ļ�¼���������µ����ݼ�
'    0117. zlCheckNoSettlementMoney:����������۲����Ƿ����δ����ý��
'    0118. zlErrBalanceCheckFromPatiID:���ݲ���ID����ʵ��ݺ��ж��쳣����
'    0119. zlCheckBalanceOverFromBalanceID:���ݽ���ID���ж��Ƿ�ǰ�����Ƿ��Ѿ����ʳɹ�
'    0120. zlCheckOtherSessionDoing:���ݽ���ID����鵱ǰ�����Ƿ������Ựվ��
'  02��һ��ͨ�ӿ����
'    0201. zlGetCardFromBalanceName:���ݽ��㷽ʽ���ƣ���ȡ������
'    0202. zlGetCardFromCardType:���ݿ����ID��ȡ������
'    0203  zlGetBalanceItemFromCardObject:���ݿ����󣬻�ȡ�µĽ�����Ϣ����
'  03�������б���غ���
'    0300. zlInitBalanceGrid:��ʼ�������б���Ϣ
'    0301. zlGetBalanceItemsFromVsBalanceGrid:���ݽ������񼰹�������ID����ȡָ���������ݼ�
'    0302. zlGetBalanceItemsFromRecord:���ݽ��ʼ�¼���ݷ���ָ���Ľ������ݼ�
'    0303. zlClearBalanceFromBalanceGrid:���ݽ��㷽ʽ�����������Ϣ��
'    0304. zlCheckBalancesIsExistFromCardTypeID:���ݿ����ID����Ƿ��ڽ����б��д��ڸý������Ľ�����Ϣ
'    0304. zlCheckVsBalanceIsExsitsFromCardObject:���ݿ����󣬼��������������Ƿ����ָ���Ľ��㷽ʽ
'    0305. zlGetBalanceItemFromBalanceGrid:����ָ���л�ȡ��������Ϣ
'    0306. zlGetBalanceItemsFromCardObject:���ݿ�����,�ӽ����б��л�ȡ��صĽ�������
'    0307. zlGetCancelBalancesFromVsBalanceGrid:���ݽ����б���ȡ���ϵĽ�����Ϣ
'    0308. zlAddBalanceDataToGridFromBalanceItems:���ݽ�����Ϣ���󣬼��ص������б���
'    0309  zlLoadBalanceItemsToVsGrid:���ݽ�����Ϣ���󣬻��ص������б�
'    0310. zlGetBalanceNULLRow:��ȡ����
'    0311. zlRecalItemObjectRowNo:���ݽ����б���Ϣ,����ˢ�½��������к�����
'    0312. zlSetBalanceRowDataFromItemsObject:���ݽ�����Ϣ��������ؽ����б��е�������
'    0313. zlSetBalanceRowDataFromItemObject:���ݽ��������ָ���еĽ���״̬
'    0314. zlGetBalanceCancelSQL:��ȡ����ȡ����������SQL
'    0315. zlMoveRowBalanceFromSwapID:���ݹ�������ID,ɾ����Ӧ�Ľ����б��ж�Ӧ����
'    0317. zlReCalcBalanceInfor:���¼��������Ϣ��δ�����Ѹ�
'    0318. zlCheckMulitInterfaceNumValied:�������ͬʱ�����������Ͻӿ�(��������)
'    0319. zlGetPtBalanceItemsFromVsBalance:��ȡ��ͨ�Ľ�����Ϣ����
'    0320. zlSetVsBalanceEditStatus:���ý���ı���״̬
'    0321. zlGetLedDisplayBankDatasFromVsBalance:���ݽ����б���ȡ��ʾ��Led�ϵĽ������ݼ�
'    0322. zlGetBalanceIDFromBalanceNo-��ȡԭ����ID
'  04:Ԥ�����
'    0400. zlInitDepositGrid:��ʼ��Ԥ������
'    0401. zlGetDelDepositItemsFromVsDeposit:����Ԥ���б���ȡ�˿���Ϣ��
'    0402. zlGetThirdTransferItemsFromVsDeposit:����Ԥ�����б����ݣ�ת�ʣ�,��ȡ�����˿�ķ�̯��ϸ��Ϣ
'    0403. zlRecalcDepositMoney�����¼����Ԥ�����
'    0405. zlLoadDepositListFromBalanceID:���ݽ���ID��ȡ��Ԥ����Ϣ��Ϣ�����ص�Ԥ���б���
'    0406. zlLoadDepositListFromRecord:����Ԥ����¼������Ԥ����Ϣ���ص�Ԥ���б���
'    0407. zlGetThridTransItemsFromVsDepositAndTranItem:����ת�˽����»�ȡ��Ҫת�ʵ�����
'    0408. zlGetItemFromVsDepositRow:����Ԥ�����У���ȡ����Ľ�����Ϣ
'  05.�������
'    0501:zlAutoRecalFeeBalanceMoney���Զ�������̯���ʽ��
'    0502.zlLoadDetaiFeeToGridFromRecord:���ݷ��ü�¼���������ݼ��ص�������ϸ������
'    0503:zlLoadDetaiFeeToGridFromBalanceID:���ݽ���ID,�����ü��ص�������ϸ������
'    0504:zlLoadFeiMuFeeListToGridFromRecord:���ݷ��ü�¼���������ݼ��ص���Ŀ��ϸ������
'    0505:zlLoadFeiMuFeeListToGridFromBalanceID:���ݽ���ID,����Ŀ���ü��ص���Ŀ��ϸ������
'    0506. zlGetReadFeeDetailFromBalanceID:���ݽ���DI,��ȡ�����������ϸ�б�����
'    0507. zlGetExceptionBalanceData:��ȡ�쳣�Ľ������ݸ���ʱ�䷶Χ�Ͳ���Ա����
'  06:Items����������
'    0601. zlCopyNewItemFromBalanceItem:����һ���µ�Item����
'����:���˺�
'����:2018-05-23 14:40:18
'*********************************************************************************************************************************************
Public grs���㷽ʽ As ADODB.Recordset
Public Const g_BalanceRow_Color_Succes = &H80000011  '�ӿڵ��óɹ�:��ɫ
Public Const g_BalanceRow_Color_Valied = &HFF&       '�ӿڵ���ʧ��:��ɫ
Public Const g_BalanceRow_Color_Normal = &H80000008  '�����Ĳ鿴:��ɫ



Public Function zlCopyNewFeeData(ByVal bytOperation As Byte, ByVal rsFeeList As ADODB.Recordset, ByRef rsNewFeeList_Out As ADODB.Recordset, Optional strOwnerFeeType As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ����¼��������Ϊһ���µļ�¼��
    '���:bytOperation-0-�����Էѷ�������;1--����Ѫ�ѷ�������
    '     rsFeeList-ԭ��¼����
    '     strOwnerFeeType-�Է�����,����ö��ŷ���
    '����:rsNewFeeList_Out-�����µķ�����������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-10-29 16:08:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, varData As Variant, i As Long
    On Error GoTo errHandle
    
    If bytOperation = 0 Then '�ԷѲ���
        If varData = "" Then Set rsNewFeeList_Out = rsFeeList: zlCopyNewFeeData = True: Exit Function
        varData = Split(strOwnerFeeType, ",")
        For i = 0 To UBound(varData)
            strFilter = strFilter & " Or �շ����='" & Replace(varData(i), "'", "") & "'"
        Next
        strFilter = Mid(strFilter, 4)
        rsFeeList.Filter = strFilter
        Set rsNewFeeList_Out = zlDatabase.CopyNewRec(rsFeeList)
        rsFeeList.Filter = 0
        zlCopyNewFeeData = True: Exit Function
    End If
    
    If bytOperation = 1 Then 'Ѫ�ⲿ��
        strFilter = " �շ����='K'"
        strFilter = Mid(strFilter, 4)
        rsFeeList.Filter = strFilter
        Set rsNewFeeList_Out = zlDatabase.CopyNewRec(rsFeeList)
        rsFeeList.Filter = 0
        zlCopyNewFeeData = True: Exit Function
    End If
    Set rsNewFeeList_Out = rsFeeList
    zlCopyNewFeeData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlComparePatiNumsIsDiff(ByVal strPatiNums1 As String, ByVal strPatiNums2 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƚ�����סԺ�����Ƿ�һ��
    '���:strPatiNums1-����סԺ����1
    '     strPatiNums2-����סԺ����2
    '����:
    '����:һ�·���true,���򷵻�False
    '����:���˺�
    '����:2018-10-29 17:12:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, i As Long, j As Long, blnFind As Boolean

    On Error GoTo errHandle
    If strPatiNums1 = strPatiNums2 Then zlComparePatiNumsIsDiff = True: Exit Function '��ͬ��ֵ���϶�һ��
    
    varData = Split(strPatiNums1, ","): varTemp = Split(strPatiNums2, ",")
    If UBound(varData) <> UBound(varTemp) Then zlComparePatiNumsIsDiff = False: Exit Function  'סԺ������һ�����϶�һ��(��ʼ������һ���ģ�Ҳ�ж�Ϊ��һ��)
    
    For i = 0 To UBound(varData)
        blnFind = False
        For j = 0 To UBound(varTemp)
            If Val(varData(i)) = Val(varTemp(j)) Then
                blnFind = True: Exit For
            End If
        Next
        If Not blnFind Then zlComparePatiNumsIsDiff = False: Exit Function  'δ�ҵ����϶���һ��
    Next
    For i = 0 To UBound(varTemp)
        blnFind = False
        For j = 0 To UBound(varData)
            If Val(varTemp(i)) = Val(varData(j)) Then
                blnFind = True: Exit For
            End If
        Next
        If Not blnFind Then zlComparePatiNumsIsDiff = False: Exit Function  'δ�ҵ����϶���һ��
    Next
    zlComparePatiNumsIsDiff = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function zlGetBalanceDataErrFromBalanceID(ByVal lng����ID As Long, ByRef rsBalance_Out As ADODB.Recordset, _
    Optional blnDel As Boolean, Optional blnMoved As Boolean, Optional strTittle As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID��ȡ��ص��쳣����������Ϣ
    '���:lng����ID-����ID
    '     strCaptions-��������
    '     blnMoved-�Ƿ��������ʷ����ת��
    '����:rsBalance_Out-��ȡ�ɹ�ʱ�����صĽ�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-05-23 14:42:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    strTittle = IIf(strTittle = "", "���ݽ���ID��ȡ��ص��쳣����������Ϣ��zlGetBalanceDataErrFromBalanceID)", strTittle)
    strSQL = " " & _
    "    Select ���㷽ʽ, Sum(��Ԥ��) As ��Ԥ��, ��־, ���� " & _
    "    From (Select Decode(Mod(��¼����, 10), 1, '[��Ԥ��]', Nvl(a.���㷽ʽ, 'δ����')) As ���㷽ʽ, " & IIf(blnDel, "-1*", "") & " a.��Ԥ�� as ��Ԥ��, " & _
    "                Decode(Nvl(a.У�Ա�־, 0), 0, '��', 2, '��', '��') As ��־, Decode(Mod(��¼����, 10), 1, -1, Nvl(b.����, 0)) As ���� " & _
    "           From ����Ԥ����¼ A, ���㷽ʽ B " & _
    "           Where a.����id = [1] And a.���㷽ʽ = b.����(+)) A " & _
    "    Group By ���㷽ʽ, ��־, ���� " & _
    "    Having Sum(a.��Ԥ��) <> 0 " & _
    "    Order By ����"
    
    If blnMoved Then
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
    End If
    Set rsBalance_Out = zlDatabase.OpenSQLRecord(strSQL, strTittle, lng����ID)
    zlGetBalanceDataErrFromBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceDataFromBalanceID(ByVal lng����ID As Long, ByRef rsBalance_Out As ADODB.Recordset, _
    Optional blnDel As Boolean, Optional blnMoved As Boolean, Optional strTittle As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID��ȡ��صĽ���������Ϣ
    '���:lng����ID-����ID
    '     strCaptions-��������
    '     blnMoved-�Ƿ��������ʷ����ת��
    '����:rsBalance_Out-��ȡ�ɹ�ʱ�����صĽ�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-05-23 14:42:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    strTittle = IIf(strTittle = "", "���ݽ���ID��ȡ��صĽ���������Ϣ��zlGetBalanceDataFromBalanceID)", strTittle)
    
    strSQL = _
    " Select Decode(Mod(��¼����, 10), 1,'��Ԥ��',decode(���㷽ʽ,NULL,'δ��','����')) as ����,NO as ���ݺ�," & IIf(blnDel, "-1*", "") & "��Ԥ�� as ���," & _
    "       ���㷽ʽ,�������,�Ƿ����Ʊ��  " & _
    " From ����Ԥ����¼  " & _
    " Where ����ID=[1] And ��Ԥ�� <> 0 " & _
    " Order by ���� Desc,NO Desc,���㷽ʽ"
    If blnMoved Then
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
    End If
    
    Set rsBalance_Out = zlDatabase.OpenSQLRecord(strSQL, strTittle, lng����ID)
    zlGetBalanceDataFromBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetSaveThreeDelSwapBatchSQL(ByVal objItem As clsBalanceItem, ByRef cllPro As Collection, _
     ByRef objItems_Out As clsBalanceItems, ByRef strDepsoitIDs As String, Optional ByVal blnRetrunXML As Boolean, Optional ByRef strInXml_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ԭ������Ϣ��ȡ���������˿���Ϣ��SQL
    '���:objItem-��Ҫ�����˿������������Ϣ
    '     blnRetrunXML-�Ƿ񷵻�XML��
    '����:cllPro-���ص�SQL��
    '     strInXml_Out-blnRetrunXML=trueʱ������,��ʽΪ:
    '     objItems_Out-���ؽ����˿���Ϣ��ϸ
    '     strBalanceDepsoitIDs-Ԥ��ID��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-05-21 11:02:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    Dim objItemTemp As clsBalanceItem
    Dim i As Long
    
    On Error GoTo errHandle
    
    strInXml_Out = "": strDepsoitIDs = ""
    ' ����,������ˮ��,����˵��,���,Ԥ��ID
    varData = Split(objItem.Tag, "|")
    Set objItems_Out = New clsBalanceItems
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",,,,", ",")
        Set objItemTemp = New clsBalanceItem
        With objItemTemp
            Set .objCard = objItem.objCard
            .���� = varTemp(0)
            .������ˮ�� = varTemp(1)
            .����˵�� = varTemp(2)
            .������ = RoundEx(-1 * Val(varTemp(3)), 2)
            .����ID = objItem.����ID
            .�к� = objItem.�к�
            .����ʱ�� = objItem.����ʱ��
            .���㷽ʽ = objItem.���㷽ʽ
            .�����ID = objItem.�����ID
            .������� = objItem.�������
            .�Ƿ��˿�ֽ��� = objItem.�Ƿ��˿�ֽ���
            .�Ƿ�ת�� = objItem.�Ƿ�ת��
            .У�Ա�־ = objItem.У�Ա�־
            .Ԥ��ID = Val(varTemp(4))
        End With
        
        objItems_Out.AddItem objItemTemp
        If blnRetrunXML Then
            strInXml_Out = strInXml_Out & "<JS>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <KH>" & objItemTemp.���� & "</KH>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <JYLSH>" & TruncStringEx(objItemTemp.������ˮ��, True) & "</JYLSH>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <JYSM>" & TruncStringEx(objItemTemp.������ˮ��, True) & "</JYSM>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <ZFJE>" & objItemTemp.������ & "</ZFJE>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <JSLX>" & 1 & "</JSLX>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <ID>" & objItemTemp.Ԥ��ID & "</ID>" & vbCrLf
            strInXml_Out = strInXml_Out & "</JS>" & vbCrLf
        End If
        
        If zlGetSaveThirdSwapDelSQLFromBalanceItem(objItemTemp, True, cllPro) = False Then Exit Function
        strDepsoitIDs = strDepsoitIDs & "," & objItemTemp.Ԥ��ID
    Next i
    If strDepsoitIDs <> "" Then strDepsoitIDs = Mid(strDepsoitIDs, 2)
    zlGetSaveThreeDelSwapBatchSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetSaveThirdSwapDelSQLFromBalanceItem(ByVal objItem As clsBalanceItem, ByVal blnModify As Boolean, ByRef cllPro As Collection, Optional intУ�Ա�־ As Integer = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ�������ȡ���������˿��SQL
    '���:objItem-��ǰ���ʶ���
    '     blnModify-�Ƿ��޸�
    '     blnת��-�Ƿ�ǰ���е�ת�ʲ���
    '     intУ�Ա�־-1-�ӿ�δ�ɹ�;0-�ӿڵ��óɹ�
    '����:cllPro-���صĽ�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-05-20 18:08:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    ' Zl_�����˿���Ϣ_Insert
    strSQL = "Zl_�����˿���Ϣ_Insert("
    '  ����id_In     �����˿���Ϣ.����id%Type,
    strSQL = strSQL & "" & objItem.����ID & ","
    '  ��¼id_In     �����˿���Ϣ.��¼id%Type,
    strSQL = strSQL & "" & objItem.Ԥ��ID & ","
    '  ���_In       �����˿���Ϣ.���%Type,
    strSQL = strSQL & "" & Abs(objItem.������) & ","
    '  ����_In       �����˿���Ϣ.����%Type,
    strSQL = strSQL & "'" & objItem.���� & "',"
    '  ������ˮ��_In �����˿���Ϣ.������ˮ��%Type,
    strSQL = strSQL & "'" & objItem.�˿����ˮ�� & "',"
    '  ����˵��_In   �����˿���Ϣ.����˵��%Type,
    strSQL = strSQL & "'" & objItem.�˿��˵�� & "',"
    '  ��������_In   Number := 0,
    strSQL = strSQL & "'" & IIf(blnModify, 1, 0) & "',"
    '  �Ƿ�δ��_In   �����˿���Ϣ.�Ƿ�δ��%Type := 0
    strSQL = strSQL & "'" & IIf(intУ�Ա�־ = 1, 1, 0) & "',"
    '  �Ƿ�ת��_In   �����˿���Ϣ.�Ƿ�ת��%Type := 0
    strSQL = strSQL & "'" & IIf(objItem.�Ƿ�ת��, 1, 0) & "',"
    '  �����id_In   �����˿���Ϣ.�����id%Type := Null
    strSQL = strSQL & "" & IIf(objItem.�����ID = 0, "NULL", objItem.�����ID) & ","
    '  ԭ������ˮ��_In �����˿���Ϣ.ԭ������ˮ��%Type := Null,
    strSQL = strSQL & "'" & objItem.������ˮ�� & "',"
    '  ԭ����˵��_In   �����˿���Ϣ.ԭ����˵��%Type := Null
    strSQL = strSQL & "'" & objItem.����˵�� & "')"
  
    zlAddArray cllPro, strSQL
    zlGetSaveThirdSwapDelSQLFromBalanceItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetSaveThirdSwapDelSQLFromBalanceItems(ByVal objItems As clsBalanceItems, ByVal blnModify As Boolean, _
    ByRef cllPro As Collection, Optional intУ�Ա�־ As Integer = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ�������ȡ���������˿��SQL
    '���:objItem-��ǰ���ʶ���
    '     blnModify-�Ƿ��޸�
    '����:cllPro-���صĽ�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-05-20 18:08:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim objItem As clsBalanceItem
    
    On Error GoTo errHandle
    For Each objItem In objItems
        If zlGetSaveThirdSwapDelSQLFromBalanceItem(objItem, blnModify, cllPro, intУ�Ա�־) = False Then Exit Function
    Next
    zlGetSaveThirdSwapDelSQLFromBalanceItems = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlThirdDelSwapIsExsistFromBalanceID(ByVal lng����ID As Long, Optional bln��δ�� As Boolean = True, Optional strTittle As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID���ж������˿���Ƿ����
    '���:bln��δ��-�Ƿ����δ�˲��ֵļ��
    '����:
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2018-05-25 09:55:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strTittle = IIf(strTittle = "", "���ݽ���ID���ж������˿���Ƿ���ڣ�zlThirdDelSwapIsExsistFromBalanceID)", strTittle)
 
    strWhere = ""
    If Not bln��δ�� Then strWhere = " And nvl(�Ƿ�δ��,0)<>1"
    strSQL = "Select 1 From �����˿���Ϣ Where ����ID=[1] And Rownum<2 " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, strTittle, lng����ID)
    zlThirdDelSwapIsExsistFromBalanceID = Not rsTemp.EOF
    Set rsTemp = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlGetThridTransItemsFromVsDepositAndTranItem(ByVal vsDeposit As VSFlexGrid, ByVal objBalanceInfor As clsBalanceInfo, ByVal strNotCardTypeIDs As String, ByVal objTranItem As clsBalanceItem, _
    ByRef objItems_Out As clsBalanceItems, Optional dblTransMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ���б�����ȡת����
    '���:objTranItem-ת����
    '     dblTransMoney-ת�ʽ��:0��ʾ��ȡ��ͬ����������Ϣ
    '     strNotCardTypeIDs-�������Ŀ����,����ö��ŷ���
    '����:objItems_Out-�����������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-27 11:59:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, bln���ѿ� As Boolean, lngCardTypeID As Long, lngCurCardTypeID As Long
    Dim objItem As clsBalanceItem, objCard As Card
    Dim dblTranTotal As Double, dbl��Ԥ�� As Double, dblMoney As Double
    Dim blnAll As Boolean, i As Long, blnת�� As Boolean
    Dim strCardOwerCardtypeIDs As String
    
    On Error GoTo errHandle
    lngCurCardTypeID = 0
    If Not objTranItem Is Nothing Then
        Set objCard = objTranItem.objCard
        lngCurCardTypeID = objTranItem.�����ID
    End If
    
    blnAll = dblTransMoney = 0
    dblTranTotal = dblTransMoney
    If objItems_Out Is Nothing Then Set objItems_Out = New clsBalanceItems
    For i = 1 To objItems_Out.Count
        If InStr("," & strCardOwerCardtypeIDs & ",", "," & objItems_Out(i).�����ID & ",") = 0 Then
            strCardOwerCardtypeIDs = strCardOwerCardtypeIDs & "," & objItems_Out(i).�����ID
        End If
    Next
    
    With vsDeposit
        For i = .Rows - 1 To 1 Step -1
            
            lngCardTypeID = Val(.TextMatrix(i, .ColIndex("�����ID")))
 
            strNO = Trim(.TextMatrix(i, .ColIndex("���ݺ�")))
            bln���ѿ� = Val(.TextMatrix(i, .ColIndex("�Ƿ����ѿ�"))) = 1
            dbl��Ԥ�� = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
 
            blnת�� = Val(.TextMatrix(i, .ColIndex("�Ƿ�ת�ʼ�����"))) = 1
            If dblTransMoney = 0 And Not blnAll Then zlGetThridTransItemsFromVsDepositAndTranItem = True: Exit Function
            
            'Ҫ֧��ת�ʵ������������ܺϲ�ת��
            If ((lngCurCardTypeID = lngCardTypeID And bln���ѿ� = False) Or lngCurCardTypeID = 0 Or (blnת�� And lngCardTypeID <> 0 And lngCurCardTypeID = 0 And bln���ѿ� = False)) _
                And InStr("," & strNotCardTypeIDs & strCardOwerCardtypeIDs & ",", "," & lngCardTypeID & ",") = 0 Then
                
                If dblTransMoney > dbl��Ԥ�� Or blnAll Then
                    dblMoney = dbl��Ԥ��
                    If Not blnAll Then dblTransMoney = RoundEx(dblTransMoney - dbl��Ԥ��, 6)
                Else
                    dblMoney = dblTransMoney
                    dblTransMoney = 0
                End If
                
                If lngCurCardTypeID <> 0 Then
                    Set objItem = zlCopyNewItemFromBalanceItem(objTranItem)
                    If objCard Is Nothing Then Set objCard = zlGetCardFromCardType(lngCardTypeID, False, Trim(.TextMatrix(i, .ColIndex("���㷽ʽ"))))
                Else
                    Set objItem = New clsBalanceItem
                    Set objCard = zlGetCardFromCardType(lngCardTypeID, False, Trim(.TextMatrix(i, .ColIndex("���㷽ʽ"))))
                End If
                Set objItem.objCard = objCard
                objItem.���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
                objItem.��������ID = Val(.TextMatrix(i, .ColIndex("��������ID")))
                objItem.�����ID = lngCardTypeID
                objItem.���� = Trim(.TextMatrix(i, .ColIndex("����")))
                objItem.������ˮ�� = Trim(.TextMatrix(i, .ColIndex("������ˮ��")))
                objItem.����˵�� = Trim(.TextMatrix(i, .ColIndex("����˵��")))
                objItem.������� = Trim(.TextMatrix(i, .ColIndex("�������")))
                objItem.������ = RoundEx(-1 * dblMoney, 6)
                objItem.����ժҪ = Trim(.TextMatrix(i, .ColIndex("ժҪ")))
                objItem.������� = IIf(objBalanceInfor.�������� = 1, True, False)
                objItem.����ID = objBalanceInfor.����ID
                objItem.����IDs = objBalanceInfor.����ID
                objItem.����ID = objBalanceInfor.����ID
                objItem.����ʱ�� = objBalanceInfor.����ʱ��
                objItem.�������� = objCard.��������
                objItem.�Ƿ�Ԥ�� = True
                objItem.�Ƿ��˿� = True
                If lngCardTypeID <> 0 Then
                    objItem.�������� = IIf(bln���ѿ�, 5, 3) '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                ElseIf objCard.�������� = 7 Then
                    objItem.�������� = 4
                Else
                    objItem.�������� = 0
                End If
                objItem.Ԥ��ID = Val(.TextMatrix(i, .ColIndex("Ԥ��ID")))
                objItem.�Ƿ����� = objCard.�������Ĺ��� <> ""
                objItem.�Ƿ�ת�� = True
                objItem.У�Ա�־ = 1
             
                objItems_Out.AddItem objItem
                objItems_Out.������ = RoundEx(objItems_Out.������ + objItem.������, 6)
           End If
        Next
    End With
    zlGetThridTransItemsFromVsDepositAndTranItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetThirdTransferItemsFromVsDeposit(ByVal vsDeposit As VSFlexGrid, ByRef objBalanceInfor As clsBalanceInfo, ByVal objCurTranItem As clsBalanceItem, _
    ByVal objDelItems As clsBalanceItems, ByRef objTranItem_Out As clsBalanceItem) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԤ����������˿�ķ�̯��ϸ��Ϣ
    '���:vsDeposit-Ԥ����������
    '     objDelItems-��ǰ�˿���Ϣ��
    '     objTranItem-��ǰ��ת����Ŀ
    '����:objTranItem_Out-��ǰ���ص�ת����Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-06-14 15:14:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, objTempItem As clsBalanceItem
    Dim objItems As clsBalanceItems, dblMoney As Double
    Dim strDelCardTypeIDs As String, dblBalance As Double

    If objCurTranItem Is Nothing Then zlGetThirdTransferItemsFromVsDeposit = True: Exit Function
    
    '��ʽ����ת�ʽ�������
    For Each objItem In objDelItems
        strDelCardTypeIDs = strDelCardTypeIDs & "," & objItem.�����ID
    Next
    
    Err = 0: On Error GoTo errHandle
    
    If objCurTranItem Is Nothing Then Exit Function
    
    dblMoney = -1 * objCurTranItem.������
    '��һ��:�ȴ��������ת��
    If zlGetThridTransItemsFromVsDepositAndTranItem(vsDeposit, objBalanceInfor, strDelCardTypeIDs, objCurTranItem, objItems, dblMoney) = False Then Exit Function
    
    If RoundEx(dblMoney, 6) = 0 Then
        If objItems Is Nothing Then Exit Function
        If objItems.Count = 0 Then Exit Function
        Set objCurTranItem.objTag = objItems
        Set objTranItem_Out = objCurTranItem
        objTranItem_Out.������ = objCurTranItem.������
        zlGetThirdTransferItemsFromVsDeposit = True: Exit Function
    End If
    
    '�ڶ������ٴ���������ת��(ʣ�����̯)
    If zlGetThridTransItemsFromVsDepositAndTranItem(vsDeposit, objBalanceInfor, strDelCardTypeIDs, Nothing, objItems, dblMoney) = False Then Exit Function
    If objItems Is Nothing Then Exit Function
    If objItems.Count = 0 Then Exit Function
    If RoundEx(dblMoney, 6) <> 0 Then
        dblBalance = RoundEx(objBalanceInfor.��ǰ���� - objBalanceInfor.ҽ��֧���ϼ� - objBalanceInfor.����, 6)
        If RoundEx(dblBalance, 6) < 0 Then dblBalance = 0 'ҽ���������ܴ��ڷ����ܽ��
        MsgBox "�㵱ǰת�ʽ�������Ԥ���Ӧ�˵Ľ��,����!" & vbCrLf & _
               "ת�ʽ��:" & Format(-1 * objCurTranItem.������, "0.00") & vbCrLf & _
               "Ӧ�˽��:" & Format(objBalanceInfor.��Ԥ���ϼ� - dblBalance, "0.00"), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    For Each objItem In objItems
        objItem.�����ID = objCurTranItem.�����ID
    Next
    Set objCurTranItem.objTag = objItems
    Set objTranItem_Out = objCurTranItem
    objTranItem_Out.������ = objCurTranItem.������
    zlGetThirdTransferItemsFromVsDeposit = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetDelDepositItemsFromVsDeposit(ByVal objThirdSwap As clsThirdSwap, ByVal vsDeposit As VSFlexGrid, _
    ByVal dblDelTotal As Double, ByVal dblNotFeeTotal As Double, ByRef objItems_Out As clsBalanceItems, _
    Optional ByVal vsBlance As VSFlexGrid) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԤ����������˿�ķ�̯��ϸ��Ϣ
    '���:objThirdSwap-�������׽ӿڶ���
    '       vsDeposit-Ԥ����������
    '       dblNotFeeTotal-δ������ܶ�(����������ܶ�-ҽ��֧���ܶ�)
    '����:objItems_Out-��ȡ�˿���Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-06-14 15:14:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblBalanceSum As Double, dblTemp As Double, dbl��� As Double, dblԭʼ��� As Double, dbl��Ԥ�� As Double, dblMoney As Double, dblDelMoney As Double
    Dim strSwapNO As String, strSwapDemo As String, strNO As String, str���� As String, strSQL As String, str���㷽ʽ As String, strDefaultBalance As String
    Dim strErrMsg As String, strExpend As String, bln���ѿ� As Boolean, blnAdd As Boolean, bln�Ƿ�ת�� As Boolean, blnDelCash As Boolean, blnFind As Boolean
    Dim lngԤ��ID As Long, lngCardTypeID As Long, lng��� As Long, i As Long, j As Long, lng��������ID As Long, int�������� As Integer
    Dim objItemsPt As clsBalanceItems, objOldItems As clsBalanceItems
    Dim objItems As clsBalanceItems, objItemsTemp As clsBalanceItems, objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim objData As clsBalanceData, objDatasTemp As clsBalanceDatas
    Dim objDataMulit As clsBalanceDatas '������׼�
    Dim objDataSingle As clsBalanceDatas  '��һ��
    Dim objDataTrans As clsBalanceDatas 'ת�ʼ�
    Dim blnSingleDel As Boolean '�Ƿ񵥽���
    Dim cllDelSwap As Collection 'array(�����ID,�Ƿ����һ�νӿڽ���) ���Ƿ����һ�νӿڽ���:1-��;0-��)
    Dim rsTemp As ADODB.Recordset
    Dim varData As Variant, blnDelToLocalMode As Boolean  '����ʱ,��Ԥ��ʣ���ʱ,Ԥ��ʣ����Ƿ��˵�ָ�����㷽ʽ
    Dim str������� As String, str����ժҪ As String
    
    dblBalanceSum = RoundEx(dblDelTotal, 2) '���ʽ������ж�λ����ˣ�ֻ���������뵽2λ�����д���
    If dblBalanceSum <= 0 Then zlGetDelDepositItemsFromVsDeposit = True: Exit Function
    blnDelToLocalMode = gTy_System_Para.TY_Balance.blnԤ����ָ�����㷽ʽ And gTy_System_Para.TY_Balance.strԤ���˿���㷽ʽ <> ""
    '��ʼ�����ݽṹ
    Set rsTemp = New ADODB.Recordset
    rsTemp.Fields.Append "���", adInteger, , adFldIsNullable
    rsTemp.Fields.Append "����", adDouble, , adFldIsNullable    '0-������;1-�ֽ�,2-���ѿ���
    rsTemp.Fields.Append "�����ID", adVarChar, 50, adFldIsNullable
    rsTemp.Fields.Append "�Ƿ����ѿ�", adInteger, , adFldIsNullable
    rsTemp.Fields.Append "��������ID", adVarChar, 50, adFldIsNullable
    rsTemp.Fields.Append "������ˮ��", adVarChar, 100, adFldIsNullable
    rsTemp.Fields.Append "���㷽ʽ", adVarChar, 50, adFldIsNullable
    rsTemp.Fields.Append "ԭʼ���", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "��Ԥ��", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "ʣ����", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "���", adDouble, , adFldIsNullable
 
    rsTemp.CursorLocation = adUseClient
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorType = adOpenStatic
    rsTemp.Open
    
    
    On Error GoTo errHandle
    lng��� = 0
    Set cllDelSwap = New Collection
    '1.�Ȼ��ܸ�����㷽ʽ���ܽ��(ԭ���������и�������ͬһ�ʽ�����ˮ���Ѿ������δ���˿��Ҫ�ų�)
    With vsDeposit
           For i = 1 To .Rows - 1
                lngCardTypeID = Val(.TextMatrix(i, .ColIndex("�����ID")))
                strSwapNO = Trim(.TextMatrix(i, .ColIndex("������ˮ��")))
                bln���ѿ� = Val(.TextMatrix(i, .ColIndex("�Ƿ����ѿ�"))) = 1
                dbl��Ԥ�� = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                lngԤ��ID = Val(.TextMatrix(i, .ColIndex("Ԥ��ID")))
                lng��������ID = Val(.TextMatrix(i, .ColIndex("��������ID")))
                bln�Ƿ�ת�� = Val(.TextMatrix(i, .ColIndex("�Ƿ�ת�ʼ�����")))
                str���� = Val(.TextMatrix(i, .ColIndex("����")))
                dbl��� = Val(.TextMatrix(i, .ColIndex("���")))
                dblԭʼ��� = Val(.TextMatrix(i, .ColIndex("ԭʼ���")))
                strNO = .TextMatrix(i, .ColIndex("���ݺ�"))
                
                If strNO <> "" Then
                    If blnDelToLocalMode Then   'Ԥ���˿��˵�ָ�����㷽ʽ����ʼ����������Ҳ�����ӿ�
                        dblMoney = RoundEx(dblMoney + dbl��Ԥ��, 6)
                    Else
                        If strSwapNO = "" Then strSwapNO = " "
                        If lngCardTypeID <> 0 Then
                            If bln���ѿ� Then
                                rsTemp.Filter = "�����ID=" & lngCardTypeID
                            Else
                                If lng��������ID <> 0 Then
                                     rsTemp.Filter = "�����ID=" & lngCardTypeID & " and ��������ID=" & lng��������ID
                                Else
                                     rsTemp.Filter = "�����ID=" & lngCardTypeID & " and ������ˮ��='" & strSwapNO & "'"
                                End If
                            End If
                            If rsTemp.EOF Then rsTemp.AddNew: lng��� = lng��� + 1: rsTemp!��� = lng���
                        Else
                            '���������ģ�ֱ�Ӵ���
                            lng��� = lng��� + 1
                            If lng��������ID <> 0 Then
                                rsTemp.Filter = "�����ID=" & lngCardTypeID & " and ��������ID=" & lng��������ID
                            Else
                                rsTemp.Filter = "�����ID=" & lngCardTypeID & " and �Ƿ����ѿ�=" & IIf(bln���ѿ�, 1, 0) & " and ���㷽ʽ='" & Trim(.TextMatrix(i, .ColIndex("���㷽ʽ"))) & "'"
                            End If
                            If rsTemp.EOF Then rsTemp.AddNew: lng��� = lng��� + 1: rsTemp!��� = lng���
                        End If
                        rsTemp!���� = IIf(lngCardTypeID <> 0, IIf(bln���ѿ�, 2, 0), 1)
                        rsTemp!�����ID = lngCardTypeID
                        rsTemp!��������ID = lng��������ID
                        rsTemp!���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
                        rsTemp!ʣ���� = NVL(rsTemp!ʣ����, 0) + dbl��Ԥ��
                        rsTemp!���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
                        rsTemp!��� = Val(NVL(rsTemp!���, 0)) + dbl���
                        rsTemp!�Ƿ����ѿ� = IIf(bln���ѿ�, 1, 0)
                        rsTemp!ԭʼ��� = Val(NVL(rsTemp!ԭʼ���)) + IIf(dblԭʼ��� > 0, dblԭʼ���, Val(NVL(rsTemp!ԭʼ���)))
                        
                        rsTemp.Update
                    End If
                End If
           Next
    End With
    
    If blnDelToLocalMode Then   'Ԥ���˿��˵�ָ�����㷽ʽ����ʼ����������Ҳ�����ӿ�
        dblMoney = dblMoney - dblNotFeeTotal
        Call zlAddfinancialTrancsToBalanceList(vsBlance, -1 * dblMoney)
        zlGetDelDepositItemsFromVsDeposit = True
        Exit Function
    End If
    
    Set objDataMulit = New clsBalanceDatas '������׼�
    Set objDataSingle = New clsBalanceDatas '��һ��
    Set objDataTrans = New clsBalanceDatas 'ת�ʼ�
    
    Set objItemsPt = New clsBalanceItems
    Set objItems = New clsBalanceItems
    Set objOldItems = New clsBalanceItems
    dblMoney = 0
    '���������˿���Ϣ��ת����Ϣ
     With vsDeposit
        For i = .Rows - 1 To 1 Step -1
            If dblBalanceSum > 0 Then
                
                lngԤ��ID = Val(.TextMatrix(i, .ColIndex("Ԥ��ID")))
                dblMoney = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                lngCardTypeID = Val(.TextMatrix(i, .ColIndex("�����ID")))
                lng��������ID = Val(.TextMatrix(i, .ColIndex("��������ID")))
                bln���ѿ� = Val(.TextMatrix(i, .ColIndex("�Ƿ����ѿ�"))) = 1
                int�������� = Val(.TextMatrix(i, .ColIndex("��������")))
                str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
                bln�Ƿ�ת�� = Val(.TextMatrix(i, .ColIndex("�Ƿ�ת�ʼ�����"))) = 1
                strSwapNO = Trim(.TextMatrix(i, .ColIndex("������ˮ��")))
                
                If dblMoney < 0 Then GoTo GoNext   '�˿�ʱ��ֱ�Ӻ���
                
                If lngCardTypeID <> 0 Then
                   If bln���ѿ� Then
                       rsTemp.Filter = "�����id=" & lngCardTypeID & " And �Ƿ����ѿ�=1 And ʣ����>0 "
                       If rsTemp.EOF Then GoTo GoNext
                       
                       dblTemp = Val(NVL(rsTemp!ʣ����))
                       If dblTemp = 0 Then GoTo GoNext
                       
                       If dblTemp < dblMoney Then dblMoney = dblTemp '����ֻ�ܳ�ʣ����
                       
                       If dblBalanceSum > dblMoney Then
                           dblBalanceSum = RoundEx(dblBalanceSum - dblMoney, 6)
                           dblDelMoney = dblMoney
                       Else
                           'If objItem.objCard.�Ƿ�ȫ�� Then dblDelMoney = Val(NVL(rsTemp!ʣ����))    '����ȫ��
                           dblDelMoney = dblBalanceSum: dblBalanceSum = 0
                       End If
                       If dblDelMoney = 0 Then GoTo GoNext
                       
                
                       dblԭʼ��� = Val(NVL(rsTemp!ԭʼ���))
                       
                       Set objItem = New clsBalanceItem
                       With objItem
                           Set .objCard = zlGetCardFromCardType(lngCardTypeID, bln���ѿ�, str���㷽ʽ)
                           .�Ƿ�ת�� = False
                           .�������� = int��������
                           .����IDs = ""
                           .������ˮ�� = strSwapNO
                           .����˵�� = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("����˵��")))
                           .���� = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("����")))
                           .��������ID = 0
                           .������ = -1 * dblDelMoney
                           .���㷽ʽ = str���㷽ʽ
                           .������� = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("�������")))
                           .����ժҪ = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("ժҪ")))
                           .�������� = 5    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                           .�Ƿ�Ԥ�� = True
                           .�Ƿ��˿� = True
                           .�Ƿ��������� = .objCard.�Ƿ�����
                           .�Ƿ�����༭ = False
                           .�Ƿ�����ɾ�� = .�Ƿ���������
                           .δ�˽�� = Val(NVL(rsTemp!���))
                           .ԭʼ��� = dblԭʼ���
                           .�����ID = lngCardTypeID
                           .���ѿ� = bln���ѿ�
                           .�Ƿ����� = IIf(.objCard.�������Ĺ��� <> "", True, False)
                           .����ʱ�� = CDate(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("�տ�����")))
                           .Ԥ��ID = lngԤ��ID
                       End With
                       
                       Set objData = New clsBalanceData
                       Set objItemsTemp = New clsBalanceItems
                       objItemsTemp.AddItem objItem
                       objItemsTemp.������ = RoundEx(objItemsTemp.������ + objItem.������, 6)
                       objItemsTemp.�շ����� = 1
                       objItemsTemp.�Ƿ�ת�� = False
                       objData.Key = "K" & lngCardTypeID & "_1"
                       Set objData.objBalanceItems = objItemsTemp
                       objDataSingle.AddItem objData
                   Else
                       '������
                       If lng��������ID <> 0 Then
                           rsTemp.Filter = "�����id=" & lngCardTypeID & " And ��������ID=" & lng��������ID & " And ʣ����>0 "
                       Else
                            If strSwapNO = "" Then strSwapNO = " "
                            rsTemp.Filter = "�����ID=" & lngCardTypeID & " and ������ˮ��='" & strSwapNO & "' And ʣ����>0"
                       End If
                       
                       If rsTemp.EOF Then GoTo GoNext
                       dblTemp = Val(NVL(rsTemp!ʣ����))
                       If dblTemp = 0 Then GoTo GoNext
                       
                       If dblTemp < dblMoney Then dblMoney = dblTemp '����ֻ�ܳ�ʣ����
                       
                       If dblBalanceSum > dblMoney Then
                           dblBalanceSum = RoundEx(dblBalanceSum - dblMoney, 6)
                           dblDelMoney = dblMoney
                       Else
                           'If objItem.objCard.�Ƿ�ȫ�� Then dblDelMoney = Val(NVL(rsTemp!ʣ����))    '����ȫ��
                           dblDelMoney = dblBalanceSum: dblBalanceSum = 0
                       End If
                       
                       If dblDelMoney = 0 Then GoTo GoNext
                       
                
                       dblԭʼ��� = Val(NVL(rsTemp!ԭʼ���))
                       
                       Set objItem = New clsBalanceItem
                       With objItem
                           Set .objCard = zlGetCardFromCardType(lngCardTypeID, bln���ѿ�, str���㷽ʽ)
                           .objCard.�Ƿ�ת�ʼ����� = bln�Ƿ�ת��
                           .�Ƿ�ת�� = bln�Ƿ�ת��
                           .�������� = int��������
                           .����IDs = ""
                           .������ˮ�� = strSwapNO
                           .����˵�� = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("����˵��")))
                           .���� = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("����")))
                           .��������ID = lng��������ID
                           .������ = -1 * dblDelMoney
                           .���㷽ʽ = str���㷽ʽ
                           .������� = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("�������")))
                           .����ժҪ = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("ժҪ")))
                           .�������� = IIf(.�������� = 7, 4, IIf(Not bln���ѿ�, 3, 5)) '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                           .�Ƿ�Ԥ�� = True
                           .�Ƿ��˿� = True
                           .�Ƿ�����༭ = False
                           .�Ƿ�����ɾ�� = True
                           .δ�˽�� = Val(NVL(rsTemp!���))
                           .ԭʼ��� = dblԭʼ���
                           .�����ID = lngCardTypeID
                           .���ѿ� = bln���ѿ�
                           .�Ƿ����� = IIf(.objCard.�������Ĺ��� <> "", True, False)
                           .����ʱ�� = CDate(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("�տ�����")))
                           .�Ƿ��������� = .objCard.�Ƿ�����
                           .Ԥ��ID = lngԤ��ID
                           .���ѿ�ID = 0
                           .�Ƿ��˿�ֽ��� = True
                       End With
                       
                       Set objItemsTemp = New clsBalanceItems
                       
                       objItemsTemp.AddItem objItem
                       objItemsTemp.������ = objItem.������
                       objItemsTemp.�շ����� = 1
                       
                       blnAdd = False
                       If objItem.objCard.�Ƿ�ת�ʼ����� Then
                           'ת�ʼ����ۣ�����Ҫ���ýӿ�
                           blnAdd = True
                       Else
                           If Not objThirdSwap.zlThirdReturnCashCheck(objItem.objCard, objItemsTemp, blnDelCash, strDefaultBalance) Then
                               '1.��ֹ����
                               objItem.�Ƿ��������� = False
                               objItem.�Ƿ�ǿ������ = blnDelCash
                               objItem.�Ƿ�����ɾ�� = objItem.�Ƿ�ǿ������
                               blnAdd = True
                           Else
                               If blnDelCash = False Then  '�Ƿ�ȱʡ����
                                   '�������֣�����ɾ��
                                   objItem.�Ƿ�����༭ = False
                                   objItem.�Ƿ�����ɾ�� = True
                                   objItem.�Ƿ�ǿ������ = True
                                   objItem.�Ƿ��������� = True: blnAdd = True
                               ElseIf strDefaultBalance <> "" Then
                               
                                   blnFind = False
                                   For j = 1 To objItemsPt.Count
                                       If objItemsPt(j).���㷽ʽ = strDefaultBalance Then
                                           objItemsPt(j).������ = objItemsPt(j).������ + objItem.������
                                           objItemsPt.������ = objItemsPt.������ + objItem.������
                                           blnFind = True
                                           Exit For
                                       End If
                                   Next
                                   
                                   If Not blnFind Then
                                       Set objItem = New clsBalanceItem
                                       With objItem
                                           Set .objCard = zlGetCardFromBalanceName(strDefaultBalance)
                                           .���㷽ʽ = strDefaultBalance
                                           .������ = RoundEx(-1 * dblDelMoney, 6)
                                           .�Ƿ��˿� = True
                                           .�Ƿ�����༭ = False
                                           .�Ƿ�����ɾ�� = True
                                           .�������� = .objCard.��������
                                           .�������� = 0 '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                                           
                                           .Tag = "ָ��Ԥ���˿�"
                                       End With
                                       objItemsPt.AddItem objItem
                                       objItemsPt.������ = RoundEx(objItemsPt.������ + objItem.������, 6)
                                   End If
                                   
                               End If
                           End If
                       End If
                       
                       If blnAdd Then
                           If objItem.objCard.�Ƿ�ת�ʼ����� Then
                               'ת��
                               blnAdd = True
                               For Each objData In objDataTrans
                                   If objData.Key = "K" & lngCardTypeID Then
                                       objData.objBalanceItems.AddItem objItem
                                       objData.objBalanceItems.������ = RoundEx(objData.objBalanceItems.������ + objItem.������, 6)
                                       blnAdd = False
                                       Exit For
                                   End If
                               Next
                               If blnAdd Then  'δ�ҵ�����Ҫ����
                                   Set objData = New clsBalanceData
                                   Set objItemsTemp = New clsBalanceItems
                                   
                                   objItemsTemp.AddItem objItem
                                   objItemsTemp.������ = RoundEx(objItemsTemp.������ + objItem.������, 6)
                                   objItemsTemp.�շ����� = 1
                                   objItemsTemp.�Ƿ�ת�� = True
                                   
                                   objData.Key = "K" & lngCardTypeID
                                   Set objData.objBalanceItems = objItemsTemp
                                   objDataTrans.AddItem objData
                               End If
                               
                           Else
                               
                               blnFind = False
                               For j = 1 To cllDelSwap.Count
                                     varData = cllDelSwap(j)
                                     If Val(varData(0)) = lngCardTypeID Then
                                        blnSingleDel = Val(varData(1)) <> 1: blnFind = True: Exit For
                                     End If
                               Next
                               
                               If blnFind = False Then
                                   blnSingleDel = objThirdSwap.zlThirdSwapIsSwapNOCall(lngCardTypeID, bln���ѿ�, strErrMsg, strExpend)
                                   cllDelSwap.Add Array(lngCardTypeID, IIf(blnSingleDel, 0, 1))
                               End If
                               
                               
                               '����������
                               objItem.�Ƿ��˿�ֽ��� = blnSingleDel
                               If blnSingleDel Then
                                    Set objDatasTemp = objDataSingle
                               Else
                                   Set objDatasTemp = objDataMulit
                               End If
                               
                               blnAdd = True
                               For Each objData In objDatasTemp
                                   If objData.Key = "K" & lngCardTypeID Then
                                       objData.objBalanceItems.AddItem objItem
                                       objData.objBalanceItems.������ = RoundEx(objData.objBalanceItems.������ + objItem.������, 6)
                                       blnAdd = False
                                       Exit For
                                   End If
                               Next
                               If blnAdd Then  'δ�ҵ�����Ҫ����
                                   Set objData = New clsBalanceData
                                   Set objItemsTemp = New clsBalanceItems
                                   objItemsTemp.AddItem objItem
                                   objItemsTemp.������ = RoundEx(objItemsTemp.������ + objItem.������, 6)
                                   objItemsTemp.�շ����� = 1
                                   objItemsTemp.�Ƿ�ת�� = False
                                   objData.Key = "K" & lngCardTypeID
                                   Set objData.objBalanceItems = objItemsTemp
                                   objDatasTemp.AddItem objData
                               End If
                               
                           End If
                       End If
                    End If
                    
                ElseIf int�������� = 7 Then
                    '��һ��ͨ
                    rsTemp.Filter = "�����id=0  And ���㷽ʽ='" & str���㷽ʽ & "  And ��������ID=" & lng��������ID & " And ʣ����>0 "
                    
                    If rsTemp.EOF Then GoTo GoNext
                    
                    dblTemp = Val(NVL(rsTemp!ʣ����))
                    If dblTemp = 0 Then GoTo GoNext
                    
                    If dblTemp < dblMoney Then dblMoney = dblTemp '����ֻ�ܳ�ʣ����
                    
                    If dblBalanceSum > dblMoney Then
                        dblBalanceSum = dblBalanceSum - dblMoney
                        dblDelMoney = dblMoney
                    Else
                        dblDelMoney = dblBalanceSum: dblBalanceSum = 0
                    End If
                    
                    If dblDelMoney = 0 Then GoTo GoNext
                    
                    dblԭʼ��� = Val(NVL(rsTemp!ԭʼ���))
                    dbl��� = dbl��� + Val(NVL(rsTemp!���))
                    
                    Set objItem = New clsBalanceItem
                    With objItem
                        Set .objCard = zlGetCardFromCardType(lngCardTypeID, bln���ѿ�, str���㷽ʽ)
                        .�������� = int��������
                        .����IDs = ""
                        .������ˮ�� = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("������ˮ��")))
                        .����˵�� = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("����˵��")))
                        .���� = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("����")))
                        .��������ID = lng��������ID
                        .������ = RoundEx(-1 * dblDelMoney, 6)
                        .���㷽ʽ = str���㷽ʽ
                        .������� = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("�������")))
                        .����ժҪ = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("ժҪ")))
                        .�������� = IIf(.�������� = 7, 4, IIf(Not bln���ѿ�, 3, 5)) '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                        .�Ƿ�Ԥ�� = True
                        .�Ƿ��˿� = True
                        .�Ƿ�����༭ = False
                        .�Ƿ�����ɾ�� = True
                        .δ�˽�� = dbl���
                        .ԭʼ��� = dblԭʼ���
                        .�����ID = lngCardTypeID
                        .���ѿ� = bln���ѿ�
                        .�Ƿ����� = IIf(.objCard.�������Ĺ��� <> "", True, False)
                        .����ʱ�� = CDate(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("�տ�����")))
                        .�Ƿ��������� = .objCard.�Ƿ�����
                        .���ѿ�ID = 0
                        .�Ƿ��˿�ֽ��� = True
                    End With
                    objOldItems.AddItem objItem
                    objOldItems.�շ����� = 1
                    objOldItems.������ = RoundEx(objOldItems.������ + objItem.������, 6)
    
                Else
                    '����������Ϣ
                   If dblBalanceSum > dblMoney Then
                       dblBalanceSum = RoundEx(dblBalanceSum - dblMoney, 6)
                       dblDelMoney = dblMoney
                   Else
                       dblDelMoney = dblBalanceSum: dblBalanceSum = 0
                   End If
                   If dblDelMoney = 0 Then GoTo GoNext
                End If
            End If
GoNext:
        Next i
        
    End With
    
    Set objItems_Out = New clsBalanceItems
    
  
    '1.�ֽ��׽���
    For Each objData In objDataSingle
        Set objItems = objData.objBalanceItems
        If objItems.Count <> 0 Then
            For Each objItem In objItems
                objItems_Out.AddItem objItem
                objItems_Out.������ = RoundEx(objItems_Out.������ + objItem.������, 6)
            Next
        End If
    Next
    
    '2.һ�ν���
     For Each objData In objDataMulit
        Set objItems = objData.objBalanceItems
        If objItems.Count <> 0 Then
            Set objItem = New clsBalanceItem
            With objItem
               Set .objCard = objItems(1).objCard
                .��������ID = 0
                .������ˮ�� = ""
                .����˵�� = ""
                .����IDs = objItems(1).����IDs
                .���㷽ʽ = objItems(1).���㷽ʽ
                .�������� = objItems(1).��������
                .����ID = objItems(1).����ID
                .����ʱ�� = objItems(1).����ʱ��
                .�������� = objItems(1).��������
                .�����ID = objItems(1).�����ID
                .������� = objItems(1).�������
                .�Ƿ񱣴� = objItems(1).�Ƿ񱣴�
                .�Ƿ���� = objItems(1).�Ƿ����
                .�Ƿ����� = objItems(1).�Ƿ�����
                .�Ƿ�ǿ������ = objItems(1).�Ƿ�ǿ������
                .�Ƿ�ȱʡ = objItems(1).�Ƿ�ȱʡ
                .�Ƿ��˿� = objItems(1).�Ƿ��˿�
                .�Ƿ��˿�ֽ��� = objItems(1).�Ƿ��˿�ֽ���
                .�Ƿ�Ԥ�� = objItems(1).�Ƿ�Ԥ��
                .�Ƿ�����༭ = objItems(1).�Ƿ�����༭
                .�Ƿ�����ɾ�� = objItems(1).�Ƿ�����ɾ��
                .�Ƿ��������� = objItems(1).�Ƿ���������
                .�Ƿ�ת�� = objItems(1).�Ƿ�ת��
                .������� = objItems(1).�������
                .���ѿ� = objItems(1).���ѿ�
                .���ѿ�ID = objItems(1).���ѿ�ID
                .У�Ա�־ = objItems(1).У�Ա�־
                .Ԥ��ID = 0
                .ԭʼ��� = 0
                .�ɿ��� = 0
                .������ = 0
                .δ�˽�� = 0
                Set .objTag = objItems
            
            End With
            For Each objItemTemp In objItems
                objItem.������ = RoundEx(objItem.������ + objItemTemp.������, 6)
                objItem.ԭʼ��� = RoundEx(objItem.ԭʼ��� + objItemTemp.ԭʼ���, 6)
                objItem.δ�˽�� = RoundEx(objItem.δ�˽�� + objItemTemp.δ�˽��, 6)
                objItems.������ = RoundEx(objItems.������ + objItemTemp.������, 6)
            Next
            objItems_Out.AddItem objItem
            objItems_Out.������ = RoundEx(objItems_Out.������ + objItem.������, 6)
        End If
    Next
   '3-ת��
    For Each objData In objDataTrans
      Set objItems = objData.objBalanceItems
      If objItems.Count <> 0 Then
      
        Set objItem = New clsBalanceItem
        With objItem
           Set .objCard = objItems(1).objCard
            .��������ID = 0
            .������ˮ�� = ""
            .����˵�� = ""
            .����IDs = objItems(1).����IDs
            .���㷽ʽ = objItems(1).���㷽ʽ
            .�������� = objItems(1).��������
            .����ID = objItems(1).����ID
            .����ʱ�� = objItems(1).����ʱ��
            .�������� = objItems(1).��������
            .�����ID = objItems(1).�����ID
            .������� = objItems(1).�������
            .�Ƿ񱣴� = objItems(1).�Ƿ񱣴�
            .�Ƿ��˿�ֽ��� = objItems(1).�Ƿ��˿�ֽ���
            .�Ƿ���� = objItems(1).�Ƿ����
            .�Ƿ����� = objItems(1).�Ƿ�����
            .�Ƿ�ǿ������ = objItems(1).�Ƿ�ǿ������
            .�Ƿ�ȱʡ = objItems(1).�Ƿ�ȱʡ
            .�Ƿ��˿� = objItems(1).�Ƿ��˿�
            .�Ƿ��˿�ֽ��� = objItems(1).�Ƿ��˿�ֽ���
            .�Ƿ�Ԥ�� = objItems(1).�Ƿ�Ԥ��
            .�Ƿ�����༭ = objItems(1).�Ƿ�����༭
            .�Ƿ�����ɾ�� = objItems(1).�Ƿ�����ɾ��
            .�Ƿ��������� = objItems(1).�Ƿ���������
            .�Ƿ�ת�� = objItems(1).�Ƿ�ת��
            .������� = objItems(1).�������
            .���ѿ� = objItems(1).���ѿ�
            .���ѿ�ID = objItems(1).���ѿ�ID
            .У�Ա�־ = objItems(1).У�Ա�־
            .Ԥ��ID = 0
            .�ɿ��� = 0
            .������ = 0
            .δ�˽�� = 0
            Set .objTag = objItems
        End With
        For Each objItemTemp In objItems
            objItem.������ = RoundEx(objItem.������ + objItemTemp.������, 6)
            objItem.ԭʼ��� = RoundEx(objItem.ԭʼ��� + objItemTemp.ԭʼ���, 6)
            objItem.δ�˽�� = RoundEx(objItem.δ�˽�� + objItemTemp.δ�˽��, 6)
        Next
        objItems_Out.AddItem objItem
        objItems_Out.������ = RoundEx(objItems_Out.������ + objItem.������, 6)
      End If
    Next
    '3.��һ��ͨ
    For Each objItem In objOldItems
        objItems_Out.AddItem objItem
        objItems_Out.������ = RoundEx(objItems_Out.������ + objItem.������, 6)
    Next
    '4.��ͨ����
    For Each objItem In objItemsPt
        objItems_Out.AddItem objItem
        objItems_Out.������ = RoundEx(objItems_Out.������ + objItem.������, 6)
    Next
    zlGetDelDepositItemsFromVsDeposit = True
    
    '�ͷ���Դ
    Set objData = Nothing: Set objDatasTemp = Nothing
    Set objDataMulit = Nothing: Set objDataSingle = Nothing: Set objDataTrans = Nothing
    Set objItemsPt = Nothing: Set objOldItems = Nothing
    Set objItems = Nothing: Set objItemsTemp = Nothing: Set objItem = Nothing
    Set objItemTemp = Nothing
    Set cllDelSwap = Nothing
    Set rsTemp = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlCheckBalancesIsExistFromCardTypeID(ByVal vsBalance As VSFlexGrid, ByVal lng�����ID As Long, Optional ByVal bln���ѿ� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ����ID����Ƿ��ڽ����б��д��ڸý������Ľ�����Ϣ
    '���:lng�����ID-�����ID
    '     bln���ѿ�-���ѿ�
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-06-20 10:21:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, blnSquare As Boolean
    Dim i As Long
    On Error GoTo errHandle
    With vsBalance
        For i = 1 To .Rows - 1
            lngCardTypeID = Val(.TextMatrix(i, .ColIndex("�����ID")))
            '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            blnSquare = Val(.TextMatrix(i, .ColIndex("����"))) = 5
            If lngCardTypeID = lng�����ID And blnSquare = bln���ѿ� Then
                zlCheckBalancesIsExistFromCardTypeID = True: Exit Function
            End If
        Next
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet���㷽ʽ() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㷽ʽ
    '����:���ؽ��㷽ʽ��Ϣ��
    '����:���˺�
    '����:2018-03-29 17:35:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    If Not grs���㷽ʽ Is Nothing Then
        If grs���㷽ʽ.State = 1 Then
            grs���㷽ʽ.Filter = 0
            Set zlGet���㷽ʽ = grs���㷽ʽ: Exit Function
        End If
    End If
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select a.����,a.����, a.����,b.Ӧ�ó���,nvl(a.Ӧ����,0) as Ӧ����,nvl(a.Ӧ�տ�,0) as Ӧ�տ�,nvl(a.ȱʡ��־,0) as ȱʡ,nvl(b.ȱʡ��־,0) as  ����ȱʡ" & vbNewLine & _
    "   From ���㷽ʽ a, ���㷽ʽӦ�� b" & vbNewLine & _
    "   Where a.���� = b.���㷽ʽ(+)    "
        
    Set grs���㷽ʽ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���㷽ʽ")
    Set zlGet���㷽ʽ = grs���㷽ʽ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardFromBalanceName(ByVal str���㷽ʽ As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ��㷽ʽ��ȡ������
    '���:str���㷽ʽ-���㷽ʽ����
    '����:
    '����:���ؿ�����
    '����:���˺�
    '����:2018-03-30 15:39:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo errHandle
    Set objCard = New Card
    Set rsTemp = zlGet���㷽ʽ
    With objCard
        .���㷽ʽ = str���㷽ʽ
        .���� = str���㷽ʽ
        rsTemp.Filter = "����='" & str���㷽ʽ & "'"
        If Not rsTemp.EOF Then
            .�������� = Val(NVL(rsTemp!����))
            .ȱʡ��־ = Val(NVL(rsTemp!����ȱʡ)) = 1
        End If
    End With
    Set zlGetCardFromBalanceName = objCard
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlGetBalanceItemFromBalanceGrid(ByVal vsGrid As VSFlexGrid, ByVal lngRow As Long, ByRef objBalanceItem_Out As clsBalanceItem) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ��������е����ݣ���ȡָ���е�BalanceItem����
    '���:lngRow-ָ������
    '����:objBalanceItem-
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-03-30 15:22:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str���㷽ʽ As String, lng�����ID As Long, lng���ѿ�ID As Long
    Dim varTemp As Variant
    Dim objCard As Card
    
    On Error GoTo errHandle
    
    With vsGrid
    
        If lngRow = 0 Then lngRow = .Row
        If lngRow > .Rows - 1 Or lngRow < 1 Then Exit Function
        If UCase(TypeName(.RowData(lngRow))) = UCase("clsBalanceItem") Then
            Set objBalanceItem_Out = .RowData(lngRow)
            If Not objBalanceItem_Out Is Nothing Then zlGetBalanceItemFromBalanceGrid = True: Exit Function
        End If
        str���㷽ʽ = .TextMatrix(lngRow, .ColIndex("���㷽ʽ"))
        If str���㷽ʽ = "" Then Exit Function
        lng�����ID = Val(.TextMatrix(lngRow, .ColIndex("�����ID")))
        lng���ѿ�ID = Val(.TextMatrix(lngRow, .ColIndex("���ѿ�ID")))
        
        If lng�����ID = 0 Then
            Set objCard = zlGetCardFromBalanceName(str���㷽ʽ)
        Else
            Call gobjSquare.objOneCardComLib.zlGetCard(lng�����ID, lng���ѿ�ID <> 0, objCard)
        End If
        
        varTemp = Split(.TextMatrix(lngRow, .ColIndex("�༭״̬")) & "|", "|")
        Set objBalanceItem_Out = New clsBalanceItem
        With objBalanceItem_Out
            Set .objCard = objCard
            .��������ID = Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("��������ID")))
            
            .������ˮ�� = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("������ˮ��"))
            .����˵�� = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("����˵��"))
            .������� = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("�������"))
            .����ժҪ = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("��ע"))
            .���� = vsGrid.Cell(flexcpData, lngRow, vsGrid.ColIndex("����"))
            .�Ƿ����� = Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("�Ƿ�����"))) = 1
            .������ = Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("������")))
            
            .�Ƿ�����༭ = Val(varTemp(0)) = 1
            .�Ƿ�����ɾ�� = Val(varTemp(1)) = 1
            .������� = CStr(vsGrid.Cell(flexcpData, lngRow, vsGrid.ColIndex("�����ID")))
            .���ѿ� = lng���ѿ�ID <> 0
            .���ѿ�ID = lng���ѿ�ID
            .�����ID = lng�����ID
            .���� = ""
            .У�Ա�־ = Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("У�Ա�־")))
            .�������� = objCard.��������
        End With
       .RowData(lngRow) = objBalanceItem_Out
    End With
    zlGetBalanceItemFromBalanceGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalancePatiNums(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal blnZero As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID��ȡ���˽��ʵ���ЧסԺ����
    '���:lng����ID-����ID
    '     lng��ҳID-���һ�ε���ҳID
    '     blnZero-�Ƿ���������
    '����:
    '����:���ؽ��ʵ���Ч����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select Zl_Fun_Getbalancepatinums([1],[2],[3]) As סԺ���� From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˵���Ч����", lng����ID, lng��ҳID, IIf(blnZero, 1, 0))
    zlGetBalancePatiNums = NVL(rsTemp!סԺ����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub zlCalcMzDepsitFromMoney(ByVal vsDepositGrid As VSFlexGrid, ByRef dblMoney As Double, ByRef dblToTal_Total As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������Ԥ��
    '���:vsDepositGrid-Ԥ������
    '     dblMoney-���μ���Ľ��
    '����:dblToTal_Total-���س�Ԥ���ܶ�
    '����:���˺�
    '����:2018-12-07 15:36:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    If dblMoney < 0 Then Exit Sub
    With vsDepositGrid
        dblToTal_Total = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" Then
                If Val(.TextMatrix(i, .ColIndex("�༭״̬"))) = 0 Then
                    If .TextMatrix(i, .ColIndex("���")) = "����" Then
                          If dblMoney > 0 Then
                              If Val(.TextMatrix(i, .ColIndex("���"))) <= dblMoney Or dblMoney < 0 Then
                                  .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                              Else
                                  .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(dblMoney, "0.00")
                              End If
                              dblToTal_Total = dblToTal_Total + RoundEx(Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), 2)
                              dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                          Else
                             .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(0, "0.00")
                          End If
                    End If
                End If
            End If
            Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlRecalcDepositMoney(ByVal bytOperationType As Byte, ByVal vsDepositGrid As VSFlexGrid, _
    ByRef objBalanceDatas As clsBalanceInfo, _
    ByVal byt����Ԥ��ȱʡʹ�÷�ʽ As Byte, ByVal bln��;������Ԥ�� As Boolean, _
    Optional ByVal dblMoney As Double = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼����Ԥ�����
    '���:bytOperationType-��������(0-������г�Ԥ��;1-��ȱʡʹ��Ԥ����;2-��ָ���������Ԥ��(��ʱ���Ⱥ�����̯��;3-ȫ��;4-סԺȫ������,����ʹ������Ԥ��
    '     vsDepositGrid-Ԥ������
    '     objBalanceDatas-��ǰ�Ľ������ݼ�
    '     dblMoneny-��Ԥ�����
    '����:���˺�
    '����:2018-03-30 11:30:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytCurFun As Byte  '0-ȫ��Ԥ����;1-�����ʽ������Ԥ��;2-ʹ������Ԥ����;3-סԺȫ������,����ʹ������Ԥ��
    Dim dblTotal As Double, i As Long, dblTemp As Double
    Dim bln������� As Boolean, bln��;���� As Boolean
    Dim bln��������Ԥ�� As Boolean
    On Error GoTo errHandle
    
    If objBalanceDatas Is Nothing Then Exit Sub
    
    
    bln������� = objBalanceDatas.�������� = 1  '���������
    bln��;���� = objBalanceDatas.�Ƿ���;����
    
    Select Case bytOperationType
    Case 0  '0-������г�Ԥ��
        bytCurFun = 0
    Case 1  '1-��ȱʡʹ��Ԥ����
        bytCurFun = 1   '������ʻ���;���ʣ�ȱʡ�����ʽ����ʹ��
        If bln������� Then
            Select Case byt����Ԥ��ȱʡʹ�÷�ʽ   '����Ԥ��ȱʡʹ�÷�ʽ
            Case 0 ' 0-ȱʡ��ʹ�ý�;1-�����ʽ��ʹ��Ԥ��;2-ʹ������Ԥ��
                bytCurFun = 0
            Case 1 '1-�����ʽ��ʹ��Ԥ��
                bytCurFun = 1
            Case 2 '2-ʹ������Ԥ��
                bytCurFun = 2
            End Select
        Else    'סԺԤ��
           If Not bln��;���� Or bln��;������Ԥ�� Then
                bytCurFun = IIf(gTy_System_Para.TY_Balance.bln����ʹ������Ԥ��, 3, 2)
           End If
        End If
        dblMoney = RoundEx(objBalanceDatas.δ���ϼ�, 2)
        
    Case 2 '2-��ָ���������Ԥ��(��ʱ���Ⱥ�����̯��
        bytCurFun = 1
        If dblMoney = 0 Then dblMoney = RoundEx(objBalanceDatas.δ���ϼ�, 2)
    Case 3 '3-ȫ��
        bytCurFun = 2
    Case 4 '4-סԺȫ������,����ʹ������Ԥ��
        bytCurFun = 3
        If dblMoney = 0 Then dblMoney = RoundEx(objBalanceDatas.δ���ϼ�, 2)
    Case Else
         bytCurFun = 0
    End Select
    
    If dblMoney < 0 Then dblMoney = 0
    bln��������Ԥ�� = False
    With vsDepositGrid
        dblTotal = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" Then
                If Val(.TextMatrix(i, .ColIndex("�༭״̬"))) = 0 Then
                    .Cell(flexcpText, i, .ColIndex("��Ԥ��"), i, .ColIndex("��Ԥ��")) = "0.00"
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                    
                    Select Case bytCurFun
                        Case 1 '�����ʽ��ʹ��
                            If dblMoney <> 0 Then
                                If Val(.TextMatrix(i, .ColIndex("���"))) <= dblMoney Then
                                      .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                                Else
                                    .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(dblMoney, "0.00")
                                End If
                                dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), 2)
                                dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                            Else
                               .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(0, "0.00")
                            End If
                        Case 2 'ȫ��
                            .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                            dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), 2)
                        Case 3 'סԺȫ������,����ʹ������Ԥ��
                            If .TextMatrix(i, .ColIndex("���")) <> "����" Then
                                .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                                dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), 2)
                                dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                            Else
                                .Cell(flexcpText, i, .ColIndex("��Ԥ��"), i, .ColIndex("��Ԥ��")) = "0.00"
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                                If bln��������Ԥ�� = False Then bln��������Ԥ�� = True
                            End If
                        Case 0 '���
                            .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(0, "0.00")
                        Case Else
                    End Select
                Else
                    dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), 2)
                    dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                End If
            End If
            Next
    End With
    dblTemp = 0
    If bytCurFun = 3 And bln��������Ԥ�� Then Call zlCalcMzDepsitFromMoney(vsDepositGrid, dblMoney, dblTemp)
    objBalanceDatas.��Ԥ���ϼ� = RoundEx(dblTotal + dblTemp, 6)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlLoadDepositListFromRecord(ByVal bytRecalDepsoit As Byte, ByVal rsDeposit As ADODB.Recordset, ByVal dblδ���ϼ� As Double, vsDeposit As VSFlexGrid, _
    ByRef dblTotal_Out As Double, ByRef intCountBill_Out As Integer, ByVal lngModul As Long, _
    Optional strFormName As String, Optional strRegKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ����¼����Ԥ����Ϣ���ص�Ԥ���б���
    '���:rsDeposit-Ԥ����¼����Ϣ
    '    dblδ���ϼ�-δ���ϼ�
    '    bytRecalDepsoit-��Ԥ������״̬:0-������,1-��������;2- ȫ������;3-סԺȫ������,����ʹ������Ԥ��
    '����:dblTotal_Out-��Ԥ���ܼ�
    '     intCountBill_Out-�漰��Ʊ������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-25 14:29:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lng�����ID As Long, dblTotal As Double
    On Error GoTo errHandle
    
    intCountBill_Out = 0: dblTotal_Out = 0
    If rsDeposit Is Nothing Then Exit Function
    If rsDeposit.State <> 1 Then Exit Function
    
    If rsDeposit.RecordCount <> 0 Then rsDeposit.MoveFirst
            
    With vsDeposit
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        i = 1
        Do While Not rsDeposit.EOF
            
            .RowData(i) = Val(NVL(rsDeposit!��¼״̬))
             lng�����ID = Val(NVL(rsDeposit!�����ID))
             If lng�����ID = 0 Then lng�����ID = Val(NVL(rsDeposit!���㿨���))
            
            .TextMatrix(i, .ColIndex("ID")) = rsDeposit!ID
            .TextMatrix(i, .ColIndex("���ݺ�")) = rsDeposit!NO
            .TextMatrix(i, .ColIndex("���")) = NVL(rsDeposit!Ԥ�����)
            
            .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) = "" & rsDeposit!Ʊ�ݺ�
            .TextMatrix(i, .ColIndex("�տ�����")) = Format(rsDeposit!����, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("���㷽ʽ")) = NVL(rsDeposit!���㷽ʽ)
            .TextMatrix(i, .ColIndex("���")) = Format(rsDeposit!���, "0.00")
            .TextMatrix(i, .ColIndex("Ԥ��ID")) = NVL(rsDeposit!Ԥ��ID)
            .TextMatrix(i, .ColIndex("��������ID")) = NVL(rsDeposit!��������ID)
            .TextMatrix(i, .ColIndex("�����ID")) = lng�����ID
            .TextMatrix(i, .ColIndex("�Ƿ����ѿ�")) = Val(NVL(rsDeposit!�Ƿ����ѿ�))
            .TextMatrix(i, .ColIndex("����")) = NVL(rsDeposit!����)
            .TextMatrix(i, .ColIndex("���������")) = NVL(rsDeposit!���������)
            .TextMatrix(i, .ColIndex("������ˮ��")) = NVL(rsDeposit!������ˮ��)
            .TextMatrix(i, .ColIndex("����˵��")) = NVL(rsDeposit!����˵��)
            .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(NVL(rsDeposit!�Ƿ�����))
            .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = Val(NVL(rsDeposit!�Ƿ�ȫ��))
            .TextMatrix(i, .ColIndex("�Ƿ�ȱʡ����")) = Val(NVL(rsDeposit!�Ƿ�ȱʡ����))
            .TextMatrix(i, .ColIndex("�Ƿ�ת�ʼ�����")) = Val(NVL(rsDeposit!ת�ʼ�����))
            .TextMatrix(i, .ColIndex("��������")) = Val(NVL(rsDeposit!��������))
            .TextMatrix(i, .ColIndex("ԭʼ���")) = Val(NVL(rsDeposit!ԭʼ���))
            .TextMatrix(i, .ColIndex("�������")) = NVL(rsDeposit!�������)
            .TextMatrix(i, .ColIndex("ժҪ")) = NVL(rsDeposit!ժҪ)
            
            Select Case bytRecalDepsoit
            Case 1 '��������
                If Val(NVL(rsDeposit!���)) <= dblδ���ϼ� Then
                    .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(rsDeposit!���, "0.00")
                    dblδ���ϼ� = dblδ���ϼ� - RoundEx(Val(NVL(rsDeposit!���)), 2)
                ElseIf dblδ���ϼ� <> 0 Then
                    .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(dblδ���ϼ�, "0.00")
                    dblδ���ϼ� = 0
                End If
            Case 2 'ȫ������
                .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(rsDeposit!���, "0.00")
            Case 3 '3-סԺȫ������,����ʹ������Ԥ��
                If .TextMatrix(i, .ColIndex("���")) <> "����" Then
                    .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(rsDeposit!���, "0.00")
                    dblδ���ϼ� = dblδ���ϼ� - RoundEx(Val(NVL(rsDeposit!���)), 2)
                Else
                    If Val(NVL(rsDeposit!���)) <= dblδ���ϼ� Then
                        .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(rsDeposit!���, "0.00")
                        dblδ���ϼ� = dblδ���ϼ� - RoundEx(Val(NVL(rsDeposit!���)), 2)
                    ElseIf dblδ���ϼ� > 0 Then
                        .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(dblδ���ϼ�, "0.00")
                        dblδ���ϼ� = 0
                    End If
                End If
            Case Else '0 -������
            End Select
            
            dblTotal = dblTotal + RoundEx(Val(NVL(rsDeposit!���)), 2)
            i = i + 1: .Rows = .Rows + 1
            rsDeposit.MoveNext
        Loop
        
        .Row = 1: .Col = .Cols - 1
        If i >= 2 And .Rows >= 2 Then .Rows = .Rows - 1
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore lngModul, vsDeposit, strFormName, strRegKey
    
    If rsDeposit.RecordCount <> 0 Then rsDeposit.MoveFirst
    intCountBill_Out = rsDeposit.RecordCount
    dblTotal_Out = dblTotal
    zlLoadDepositListFromRecord = True
    Exit Function
errHandle:
    vsDeposit.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetClassMoney(ByVal lng����ID As Long, ByVal dbl��ǰ���ʽ�� As Double, ByRef rsMoney As ADODB.Recordset, _
    rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '���:lng����ID-����ID,Ϊ0ʱ,��RsFeeListΪ׼
    '     dbl��ǰ���ʽ��-��ǰ���ʽ��(��Ҫ�Ƿ�̯ʱʹ��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, dblMoney As Double
    Dim dblTemp As Double
    
    On Error GoTo errHandle
    
    '��ʼ�����ݽṹ
    Set rsMoney = New ADODB.Recordset
    rsMoney.Fields.Append "�շ����", adVarChar, 10, adFldIsNullable
    rsMoney.Fields.Append "���", adDouble, , adFldIsNullable
    rsMoney.CursorLocation = adUseClient
    rsMoney.LockType = adLockOptimistic
    rsMoney.CursorType = adOpenStatic
    rsMoney.Open
        
    If lng����ID <> 0 Then
        strSQL = "" & _
        "   Select  A.�շ����,nvl(sum(A.���ʽ��) ,0) as ���   " & _
        "   From ������ü�¼ A" & _
        "   Where A.����ID=[1] Group by A.�շ���� " & _
        "   Union ALL " & _
        "   Select  A.�շ����,nvl(sum(A.���ʽ��) ,0) as ���   " & _
        "   From סԺ���ü�¼ A" & _
        "   Where A.����ID=[1] Group by A.�շ���� "
        strSQL = "Select �շ����,Sum(���) as ��� From (" & strSQL & ")  Group by  �շ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շ������", lng����ID)
    
        With rsTemp
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                rsMoney.Find "�շ����='" & NVL(!�շ����, "��") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!�շ���� = NVL(!�շ����, "��")
                rsMoney!��� = Val(NVL(rsMoney!���)) + Val(NVL(!���))
                rsMoney.Update
                .MoveNext
            Loop
        End With
        zlGetClassMoney = True
        Exit Function
    End If
    
    If rsFeeList Is Nothing Then Exit Function
    
    With rsFeeList
        dblMoney = dbl��ǰ���ʽ��
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dblTemp = Val(NVL(!δ����))
            If RoundEx(dblMoney - dblTemp, gbytDec) <= 0 Then
                dblTemp = dblMoney
            End If
            If dblTemp <> 0 And dblMoney <> 0 Then
                rsMoney.Find "�շ����='" & NVL(!�շ����, "��") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!�շ���� = NVL(!�շ����, "��")
                rsMoney!��� = Val(NVL(rsMoney!���)) + dblTemp
                rsMoney.Update
            End If
            dblMoney = dblMoney - dblTemp
            .MoveNext
        Loop
    End With
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceItemsFromCardObject(ByVal vsGrid As VSFlexGrid, ByVal objCard As Card, ByRef objBalanceItems_out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ����󣬴ӽ����б��л�ȡ��صĽ�������
    '���:objCard-��ǰ������
    '����:objBalanceItems_Out-���ؽ������ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-03-30 15:22:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str���㷽ʽ As String, lng�����ID As Long, lng���ѿ�ID As Long
    Dim varTemp As Variant
    Dim objItem As clsBalanceItem
    
    
    On Error GoTo errHandle
    
    If objCard Is Nothing Then Exit Function
    Set objBalanceItems_out = New clsBalanceItems
    
    With vsGrid
        For i = 1 To .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, i, objItem) Then
                If (objItem.�����ID = objCard.�ӿ���� And objItem.���ѿ� = objCard.���ѿ� And (objItem.���ѿ�ID <> 0 And objItem.���ѿ� Or objItem.���ѿ� = False)) Or (objCard.�ӿ���� <= 0 And objCard.���㷽ʽ = objItem.���㷽ʽ) Then
                    objBalanceItems_out.AddItem objItem
                    objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
                End If
            End If
        Next
    End With
    zlGetBalanceItemsFromCardObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceItemsFromGrid(ByVal vsGrid As VSFlexGrid, ByVal int���� As Integer, ByRef objBalanceItems_out As clsBalanceItems, Optional objBalanceItem As clsBalanceItem, Optional bln�˿� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������񣬻�ȡ���е����ѿ�������Ϣ��
    '���:vsGrid-�������
    '     int����-0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    '     objItem-���Nothing,���ʾ������ȡ��,���򰴵�ǰ��Ŀȡ��
    '     bln�˿�-�Ƿ��ȡ�˿bln�˿�-true,����:objitem��������෵������Ҫ�Ƿ����˿�)
    '����:objBalanceItems_Out-�������ѿ������Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-04-11 18:42:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    Dim i As Long
    On Error GoTo errHandle
    Set objBalanceItems_out = New clsBalanceItems
    With vsGrid
        For i = 1 To .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, i, objItem) Then
                If Not objBalanceItem Is Nothing Then
                    If objItem.�������� = int���� And ( _
                        (objBalanceItem.��������ID = objItem.��������ID And objBalanceItem.�����ID = objItem.�����ID And objBalanceItem.���ѿ� = objItem.���ѿ�) _
                        Or (objBalanceItem.�����ID <= 0 And objItem.�����ID <= 0 And objBalanceItem.���㷽ʽ = objItem.���㷽ʽ)) Then
                        
                        Set objItem = zlCopyNewItemFromBalanceItem(objItem)
                        If bln�˿� Then
                            If objItem.ԭʼ��� = 0 Then objItem.ԭʼ��� = objItem.������
                            If objItem.δ�˽�� = 0 Then objItem.δ�˽�� = objItem.������
                            objItem.������ = RoundEx(-1 * objItem.������, 6)
                        End If
                        
                        objBalanceItems_out.AddItem objItem
                        objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
                    End If
                Else
                    If objItem.�������� = int���� Then
                        Set objItem = zlCopyNewItemFromBalanceItem(objItem)
                        If bln�˿� Then
                            If objItem.ԭʼ��� = 0 Then objItem.ԭʼ��� = objItem.������
                            If objItem.δ�˽�� = 0 Then objItem.δ�˽�� = objItem.������
                            objItem.������ = RoundEx(-1 * objItem.������, 6)
                        End If
                        objBalanceItems_out.AddItem objItem
                        objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
                    End If
                End If
            End If
        Next
    End With
    zlGetBalanceItemsFromGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetDelThirdDepositBalance(ByVal vsGrid As VSFlexGrid, ByRef objBalanceDatas_Out As clsBalanceDatas) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������Ԥ�����˿���Ϣ
    '���:
    '����:clsBalanceDatas-���ؽ�����Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-04-16 14:13:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, objData As clsBalanceData
    Dim objItems As clsBalanceItems
    Dim strCardTypeIDs As String
    Dim i As Long
    Set objBalanceDatas_Out = New clsBalanceDatas
    strCardTypeIDs = ""
    With vsGrid
        For i = 1 To .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, i, objItem) Then
                If objItem.�Ƿ�Ԥ�� And objItem.�Ƿ��˿� And objItem.�Ƿ���� = False And objItem.�����ID <> 0 And objItem.���ѿ� = False Then
                    If InStr(strCardTypeIDs & ",", "," & objItem.�����ID & ",") = 0 Then
                        Set objItems = New clsBalanceItems
                        Set objData = New clsBalanceData
                        Set objData.objBalanceItems = objItems
                        objData.Key = "K" & objItem.�����ID
                        Call objBalanceDatas_Out.AddItem(objData, "K" & objItem.�����ID)
                        strCardTypeIDs = strCardTypeIDs & "," & objItem.�����ID
                    End If
                    
                    objBalanceDatas_Out("K" & objItem.�����ID).objBalanceItems.AddItem objItem
                    With objBalanceDatas_Out("K" & objItem.�����ID).objBalanceItems
                        .������ = .������ + objItem.������
                        .�շ����� = 1
                    End With
                End If
            End If
        Next
    End With
    zlGetDelThirdDepositBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlThirdDelMoneyIsExistFromVsGrid(ByVal vsGrid As VSFlexGrid, ByVal lng�����ID As Long, ByVal lng��������ID As Long, str������ˮ�� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���Ŀ���𼰽����Ƿ��ڽ��㷽ʽ��Ϣ�д����˿�
    '���:
    '����:
    '����:�����˿��true,���򷵻�False
    '����:���˺�
    '����:2018-04-18 23:42:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    Dim i As Long
    
    On Error GoTo errHandle
    
    With vsGrid
        For i = 1 To .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, i, objItem) Then
               If objItem.�����ID = lng�����ID And (objItem.��������ID = lng��������ID Or objItem.objCard.�Ƿ�ת�ʼ�����) And objItem.������ < 0 Then
                    zlThirdDelMoneyIsExistFromVsGrid = True: Exit Function
               End If
            End If
        Next
    End With
    zlThirdDelMoneyIsExistFromVsGrid = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Public Function zlCopyNewItemFromBalanceItem(ByVal objOldItem As clsBalanceItem) As clsBalanceItem
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ���µ�Item
    '���:objOldItem-�ɵ�Item����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-04-19 14:14:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    
    
    On Error GoTo errHandle
    Set objItem = New clsBalanceItem
    If objOldItem Is Nothing Then
        Set objItem.objCard = New Card
        Set zlCopyNewItemFromBalanceItem = objItem: Exit Function
    End If
    
    With objItem
        Set .objCard = zlCopyNewCardFromCard(objOldItem.objCard)
        .Key = objOldItem.Key
        .Tag = objOldItem.Tag
        .��������ID = objOldItem.��������ID
        .������ˮ�� = objOldItem.������ˮ��
        .����˵�� = objOldItem.����˵��
        .�ɿ��� = objOldItem.�ɿ���
        .����IDs = objOldItem.����IDs
        
        .���㷽ʽ = objOldItem.���㷽ʽ
        .������� = objOldItem.�������
        .������ = objOldItem.������
        .�������� = objOldItem.��������
        .�������� = objOldItem.��������
        .����ժҪ = objOldItem.����ժҪ
        .����ID = objOldItem.����ID
        .����ʱ�� = objOldItem.����ʱ��
        .����ID = objOldItem.����ID
        
        .���� = objOldItem.����
        .������ˮ�� = objOldItem.������ˮ��
        .����˵�� = objOldItem.����˵��
        .�����ID = objOldItem.�����ID
        .���� = objOldItem.����
        .�Ƿ���� = objOldItem.�Ƿ����
        .�Ƿ����� = objOldItem.�Ƿ�����
        .�Ƿ�ȱʡ = objOldItem.�Ƿ�ȱʡ
        .�Ƿ��˿� = objOldItem.�Ƿ��˿�
        .�Ƿ�Ԥ�� = objOldItem.�Ƿ�Ԥ��
        .�Ƿ�����༭ = objOldItem.�Ƿ�����༭
        .�Ƿ�����ɾ�� = objOldItem.�Ƿ�����ɾ��
        .�Ƿ��������� = objOldItem.�Ƿ���������
        .�Ƿ񱣴� = objOldItem.�Ƿ񱣴�
        .�Ƿ��˿�ֽ��� = objOldItem.�Ƿ��˿�ֽ���
        .δ�˽�� = objOldItem.δ�˽��
        .���� = objOldItem.����
        .������� = objOldItem.�������
        .���ѿ� = objOldItem.���ѿ�
        .���ѿ�ID = objOldItem.���ѿ�ID
        .У�Ա�־ = objOldItem.У�Ա�־
        .ԭʼ��� = objOldItem.ԭʼ���
        .�ʻ���� = objOldItem.�ʻ����
        .�˿����ˮ�� = objOldItem.�˿����ˮ��
        .�˿��˵�� = objOldItem.�˿��˵��
        .�Ҳ� = objOldItem.�Ҳ�
        .�к� = objOldItem.�к�
        .Ԥ��ID = objOldItem.Ԥ��ID
        .�Ƿ��ѻ�ҽ�� = objOldItem.�Ƿ��ѻ�ҽ��
        .QRCode = objOldItem.QRCode
        Set .objTag = Nothing
    End With

    Set zlCopyNewItemFromBalanceItem = objItem
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
     Set zlCopyNewItemFromBalanceItem = objItem
End Function
Public Function zlCopyNewCardFromCard(ByVal objOldCard As Card) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ�������󣬸���Ϊ�µĿ�����
    '���:objOldCard-�ɿ�
    '����:�����µ�Card����
    '����:���˺�
    '����:2018-04-19 14:25:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Set objCard = New Card
    If objOldCard Is Nothing Then Set zlCopyNewCardFromCard = Nothing: Exit Function
    
    On Error GoTo errHandle
    With objOldCard
        objCard.��ע = .��ע
        objCard.���� = .����
        objCard.���ܼ� = .���ܼ�
        objCard.�ӿڱ��� = .�ӿڱ���
        objCard.�ӿڳ����� = .�ӿڳ�����
        objCard.�ӿ���� = .�ӿ����
        objCard.���㷽ʽ = .���㷽ʽ
        objCard.�������� = .��������
        objCard.�������Ĺ��� = .�������Ĺ���
        objCard.���ų��� = .���ų���
        objCard.�����ظ�ʹ�� = .�����ظ�ʹ��
        objCard.�ɷ����� = .�ɷ�����
        objCard.��� = .���
        objCard.������� = .�������
        objCard.������������ = .������������
        objCard.���볤�� = .���볤��
        objCard.���볤������ = .���볤������
        objCard.���� = .����
        objCard.ģ�������� = .ģ��������
        objCard.���� = .����
        objCard.ǰ׺�ı� = .ǰ׺�ı�
        objCard.ȱʡ��־ = .ȱʡ��־
        objCard.�豸�Ƿ����ûس� = .�豸�Ƿ����ûس�
        objCard.�Ƿ�ֿ����� = .�Ƿ�ֿ�����
        objCard.�Ƿ�����ʻ� = .�Ƿ�����ʻ�
        objCard.�Ƿ񷢿� = .�Ƿ񷢿�
        objCard.�Ƿ�ǽӴ�ʽ���� = .�Ƿ�ǽӴ�ʽ����
        objCard.�Ƿ�Ӵ�ʽ���� = .�Ƿ�Ӵ�ʽ����
        objCard.�Ƿ�ģ������ = .�Ƿ�ģ������
        objCard.�Ƿ�ȫ�� = .�Ƿ�ȫ��
        objCard.�Ƿ�ȱʡ���� = .�Ƿ�ȱʡ����
        objCard.�Ƿ�ɨ�� = .�Ƿ�ɨ��
        objCard.�Ƿ�ˢ�� = .�Ƿ�ˢ��
        objCard.�Ƿ��˿��鿨 = .�Ƿ��˿��鿨
        objCard.�Ƿ����� = .�Ƿ�����
        objCard.�Ƿ�д�� = .�Ƿ�д��
        objCard.�Ƿ��ϸ���� = .�Ƿ��ϸ����
        objCard.�Ƿ�֤�� = .�Ƿ�֤��
        objCard.�Ƿ��ƿ� = .�Ƿ��ƿ�
        objCard.�Ƿ�ת�ʼ����� = .�Ƿ�ת�ʼ�����
        objCard.�Ƿ��Զ���ȡ = .�Ƿ��Զ���ȡ
        objCard.�ض���Ŀ = .�ض���Ŀ
        objCard.ͼ���ʶ = .ͼ���ʶ
        objCard.ϵͳ = .ϵͳ
        objCard.���ѿ� = .���ѿ�
        objCard.֧������ = .֧������
        objCard.֧��ͼ���ʶ = .֧��ͼ���ʶ
        objCard.�Զ���ȡ��� = .�Զ���ȡ���
        objCard.���ƿ� = .���ƿ�
        objCard.�Ƿ�֧��ɨ�븶 = .�Ƿ�֧��ɨ�븶
        objCard.�Ƿ�������� = .�Ƿ��������
    End With
    Set zlCopyNewCardFromCard = objCard
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    zlCopyNewCardFromCard = objCard
End Function

Public Function zlGetCardFromCardType(ByVal lng�����ID As Long, bln���ѿ� As Boolean, ByVal str���㷽ʽ As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ����ID��ȡ������
    '���:lng�����ID-�����ID
    '     bln���ѿ�-�Ƿ����ѿ�
    '     str���㷽ʽ-���㷽ʽ
    '����:
    '����:�ɹ�������
    '����:���˺�
    '����:2018-04-02 14:29:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As New Card
    On Error GoTo errHandle
    If lng�����ID <> 0 Then
        'zlGetCard(ByVal lngCardTypeID As Long, ByVal bln���ѿ� As Boolean,  ByRef objCard As Card) As Boolean
        If gobjSquare.objOneCardComLib.zlGetCard(lng�����ID, bln���ѿ�, objCard) = False Then
            Set objCard = zlGetCardFromBalanceName(str���㷽ʽ)
        End If
    Else
        Set objCard = zlGetCardFromBalanceName(str���㷽ʽ)
    End If
    Set zlGetCardFromCardType = objCard: Exit Function

    zlGetCardFromCardType = True
    Exit Function
errHandle:
    Set objCard = zlGetCardFromBalanceName(str���㷽ʽ)
    Set zlGetCardFromCardType = objCard: Exit Function
End Function

Public Sub zlClearBalanceFromBalanceGrid(ByRef vsGrid As VSFlexGrid, ByVal strBalance As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ��㷽ʽ�����������Ϣ��
    '����:���˺�
    '����:2018-04-16 11:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, objItem As clsBalanceItem
    On Error GoTo errHandle
    With vsGrid
        j = 1
        Do While j <= .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, j, objItem) Then
                If objItem.���㷽ʽ = strBalance And objItem.�������� = 0 Then
                    .RowData(j) = ""
                    Set objItem = Nothing
                    If .Rows >= 2 Then
                        .RemoveItem j
                    Else
                        .Rows = 2
                       .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                       .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                       .RowData(1) = ""
                       j = 2
                    End If
                Else
                      j = j + 1
                End If
            Else
                j = j + 1
            End If
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Sub zlClearDelDepositBalance(ByRef vsGrid As VSFlexGrid, Optional ByRef objItems As clsBalanceItems)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ԥ��������н�����Ϣ
    '���:objItems-������ͨ������Ϣ��(��ͬ��ɾ����������Ϊ0�Ľ����¼)
    '����:���˺�
    '����:2018-04-16 11:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, objItem As clsBalanceItem
    Dim blnDel As Boolean, objItemTemp As clsBalanceItem
    
    On Error GoTo errHandle
    With vsGrid
        j = 1
        Do While j <= .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, j, objItem) Then
            
                blnDel = (objItem.�Ƿ�Ԥ�� Or objItem.�������� = 9 Or objItem.Tag = "ָ��Ԥ���˿�")
                
                If Not objItems Is Nothing And Not blnDel Then
                    '��Ҫ�ų�Ԥ�������¼���ʱ����ָ�����㷽ʽ���ٴμ���ʱδ����ָ�����㷽ʽ������ҲҪһ�����
                    For Each objItemTemp In objItems
                        If objItemTemp.�������� = 0 And objItem.���㷽ʽ = objItemTemp.���㷽ʽ And objItemTemp.�������� = objItem.�������� Then
                            '�ҵ���:
                            blnDel = True: Exit For
                        End If
                    Next
                End If
                
                If blnDel Then
                    .RowData(j) = ""
                    Set objItem = Nothing
                    If .Rows >= 2 Then
                        .RemoveItem j
                    Else
                        .Rows = 2
                       .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                       .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                       .RowData(1) = ""
                       j = 2
                    End If
                Else
                      j = j + 1
                End If
            Else
                j = j + 1
            End If
        Loop
        If .Rows <= 1 Then .Rows = 2
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlGetPtBalanceItemsFromVsBalance(ByVal vsBalance As VSFlexGrid, objPtItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ͨ�Ľ�����Ϣ��
    '���:vsBalance-����ؼ�
    '����:objPtItems_Out-��ͨ������Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-29 12:10:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As clsBalanceItem
    
    On Error GoTo errHandle
    Set objPtItems_Out = New clsBalanceItems
    With vsBalance
        For i = 1 To .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsBalance, i, objItem) Then
                '��Ҫ�ų�Ԥ�������¼���ʱ����ָ�����㷽ʽ���ٴμ���ʱδ����ָ�����㷽ʽ������ҲҪһ�����
                If objItem.�������� = 0 And objItem.�������� <> 9 Then
                    objPtItems_Out.AddItem objItem
                    objPtItems_Out.������ = RoundEx(objPtItems_Out.������ + objItem.������, 6)
                End If
            End If
        Next
    End With
    zlGetPtBalanceItemsFromVsBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlClearBalanceFromItems(ByVal vsBalance As VSFlexGrid, objCurItems As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ͨ�Ľ�����Ϣ��
    '���:vsBalance-����ؼ�
    '����:objPtItems_Out-��ͨ������Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-29 12:10:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim j As Long, objItem As clsBalanceItem
   Dim blnDel As Boolean, objItemTemp As clsBalanceItem

    
    On Error GoTo errHandle
      
    If objCurItems Is Nothing Then Exit Function
    With vsBalance
        j = 1
        Do While j <= .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsBalance, j, objItem) Then
                '��Ҫ�ų�Ԥ�������¼���ʱ����ָ�����㷽ʽ���ٴμ���ʱδ����ָ�����㷽ʽ������ҲҪһ�����
                blnDel = False
                For Each objItemTemp In objCurItems
                    If objItem.�����ID = objItemTemp.�����ID And objItemTemp.�������� = objItem.�������� And objItemTemp.���ѿ� = objItem.���ѿ� And objItemTemp.��������ID = objItem.��������ID Then
                        '�ҵ���:
                        blnDel = True: Exit For
                    End If
                Next
                If blnDel Then
                    .RowData(j) = ""
                    Set objItem = Nothing
                    If .Rows >= 2 Then
                        .RemoveItem j
                    Else
                        .Rows = 2
                       .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                       .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                       .RowData(1) = ""
                       j = 2
                    End If
                Else
                      j = j + 1
                End If
            Else
                j = j + 1
            End If
        Loop
        If .Rows <= 1 Then .Rows = 2
    End With
    zlClearBalanceFromItems = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub zlClearDelDepositBalanceFromItems(ByRef vsGrid As VSFlexGrid, ByVal objItems As clsBalanceItems)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ԥ��������н�����Ϣ(ֻ�����ͨ������Ϣ)
    '���:objItems-������ͨ������Ϣ��(��ͬ��ɾ����������Ϊ0�Ľ����¼)
    '����:���˺�
    '����:2018-04-16 11:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, objItem As clsBalanceItem
    Dim blnDel As Boolean, objItemTemp As clsBalanceItem
    
    On Error GoTo errHandle
    If objItems Is Nothing Then Exit Sub
    
    With vsGrid
        j = 1
        Do While j <= .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, j, objItem) Then
             
                '��Ҫ�ų�Ԥ�������¼���ʱ����ָ�����㷽ʽ���ٴμ���ʱδ����ָ�����㷽ʽ������ҲҪһ�����
                blnDel = False
                For Each objItemTemp In objItems
                    If objItemTemp.�������� = 0 And objItem.���㷽ʽ = objItemTemp.���㷽ʽ And objItemTemp.�������� = objItem.�������� Then
                        '�ҵ���:
                        blnDel = True: Exit For
                    End If
                Next
                If blnDel Then
                    .RowData(j) = ""
                    Set objItem = Nothing
                    If .Rows >= 2 Then
                        .RemoveItem j
                    Else
                        .Rows = 2
                       .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                       .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                       .RowData(1) = ""
                       j = 2
                    End If
                Else
                      j = j + 1
                End If
            Else
                j = j + 1
            End If
        Loop
        If .Rows <= 1 Then .Rows = 2
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlGetThirdDelRecordFromBalanceID(ByVal lng����ID As Long, ByRef rsThirdDel_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID,��ȡ�����˿���Ϣ��
    '���:lng����ID-����ID
    '����:rsThirdDel_Out-���������˿���Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-17 15:45:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    
    strSQL = " " & _
    "   Select a.����id, a.��¼id As Ԥ��id,nvl(a.�����ID,b.�����ID ) as �����ID,a.����, a.���,b.������ˮ��, b.����˵��, a.�Ƿ�δ��, a.�Ƿ�ת��, b.��������id,b.�������,B.ժҪ,b.��� as ԭʼ���" & vbCrLf & _
    "   From �����˿���Ϣ A, ����Ԥ����¼ B " & vbCrLf & _
    "   Where a.����id =[1] And a.��¼id = b.Id" & vbCrLf & _
    " "
    Set rsThirdDel_Out = zlDatabase.OpenSQLRecord(strSQL, "���ݽ���ID��ȡ�����˿���Ϣ��ϸ", lng����ID)
    
    zlGetThirdDelRecordFromBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetDelBalanceItemsFromRecord(ByVal objCurItem As clsBalanceItem, _
    ByVal rsThirdDelRecord As ADODB.Recordset, ByRef objDelItems_Out As clsBalanceItems, _
    Optional ByVal bln�˿�ֽ��� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����˿��¼����ȡ�����˿���Ϣ��ϸ��
    '���:objCurItem-��ǰ�Ľ�����Ϣ
    '     rsThirdDelRecord-�����˿���Ϣ��
    '     bln�˿�ֽ���-�˿��Ƿ�ֽ���
    '����:objDelItems_out-���������˿���Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-17 16:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    Dim objCard As Card
    On Error GoTo errHandle
    Set objDelItems_Out = New clsBalanceItems
    If rsThirdDelRecord Is Nothing Then Exit Function
    If rsThirdDelRecord.State <> 1 Then Exit Function
    If bln�˿�ֽ��� Then
        rsThirdDelRecord.Filter = "�����ID=" & objCurItem.�����ID & " And Ԥ��ID=" & objCurItem.Ԥ��ID
    Else
        rsThirdDelRecord.Filter = "�����ID=" & objCurItem.�����ID
    End If
    With rsThirdDelRecord
        Set objCard = zlGetCardFromCardType(objCurItem.�����ID, False, "")
        Do While Not .EOF
            Set objItem = New clsBalanceItem
            Set objItem.objCard = objCard
            objItem.��������ID = Val(NVL(!��������ID))
            objItem.Ԥ��ID = Val(NVL(!Ԥ��ID))
            objItem.У�Ա�־ = IIf(Val(NVL(!�Ƿ�δ��)) = 1, 1, 2)
            objItem.�к� = objCurItem.�к�
            objItem.������ˮ�� = Trim(NVL(!������ˮ��))
            objItem.����˵�� = Trim(NVL(!����˵��))
            objItem.�ɿ��� = 0
            objItem.����IDs = objCurItem.����IDs
            objItem.���㷽ʽ = objCurItem.���㷽ʽ
            objItem.������� = Trim(NVL(!�������))
            objItem.������ = RoundEx(-1 * Val(Trim(NVL(!���))), 6)
            objItem.�������� = objCurItem.��������
            objItem.�������� = objCurItem.��������
            objItem.����ժҪ = Trim(NVL(!ժҪ))
            objItem.����ID = objCurItem.����ID
            objItem.����ʱ�� = objCurItem.����ʱ��
            objItem.���� = Trim(NVL(!����))
            objItem.�����ID = objCurItem.�����ID
            objItem.������� = objCurItem.�������
            objItem.���� = objCurItem.����
            objItem.ʣ���� = 0
            objItem.�Ƿ񱣴� = objCurItem.�Ƿ񱣴�
            objItem.�Ƿ���� = objCurItem.У�Ա�־ = 2
            objItem.�Ƿ����� = objCurItem.�Ƿ�����
            objItem.�Ƿ�ǿ������ = objCurItem.�Ƿ�ǿ������
            objItem.�Ƿ�ȱʡ = objCurItem.�Ƿ�ȱʡ
            objItem.�Ƿ��˿� = objCurItem.�Ƿ��˿�
            objItem.�Ƿ��˿�ֽ��� = objCurItem.�Ƿ��˿�ֽ���
            objItem.�Ƿ�Ԥ�� = objCurItem.�Ƿ�Ԥ��
            objItem.�Ƿ�����༭ = objCurItem.�Ƿ�����༭
            objItem.�Ƿ�����ɾ�� = objCurItem.�Ƿ�����ɾ��
            objItem.�Ƿ��������� = objCurItem.�Ƿ���������
            objItem.�Ƿ�ת�� = objCard.�Ƿ�ת�ʼ�����
            If Val(NVL(!�Ƿ�ת��)) = 1 Then objItem.�Ƿ�ת�� = True: objItem.objCard.�Ƿ�ת�ʼ����� = True
            objItem.δ�˽�� = 0
            objItem.���� = 0
            objItem.������� = objCurItem.�������
            objItem.���ѿ� = objCurItem.���ѿ�
            objItem.���ѿ�ID = objCurItem.���ѿ�ID
            objItem.ԭʼ��� = Val(Trim(NVL(!ԭʼ���)))
            objItem.�ʻ���� = 0
            objItem.�Ҳ� = 0
            If objItem.�Ƿ�ת�� Then objCurItem.�Ƿ�ת�� = True
            objDelItems_Out.AddItem objItem
            objDelItems_Out.������ = objDelItems_Out.������ + objItem.������
            objDelItems_Out.���� = objItem.��������
            objDelItems_Out.�Ƿ�ת�� = objItem.�Ƿ�ת��
            objDelItems_Out.�˷ѽ���IDs = objCurItem.����IDs
            .MoveNext
        Loop
    End With
   zlGetDelBalanceItemsFromRecord = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If


End Function
Public Function zlGetBalanceItemsFromRecord(ByVal lng����ID As Long, ByVal bytMCMode As Byte, intInsure As Integer, ByVal objThirdSwap As clsThirdSwap, _
    ByVal bln���� As Boolean, ByVal rsBalanceRecord As ADODB.Recordset, ByRef objBalanceInfor As clsBalanceInfo, ByRef objBalanceItems_out As clsBalanceItems, _
    Optional strErrMsg_out As String, Optional bytCurType As Byte, Optional ByRef blnҽ����������_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ��ʼ�¼���ݷ���ָ���Ľ������ݼ�
    '���:objCard-������
    '     lng��������Id-��������ID
    '     bytCurType-��ǰ����������:0-����;1-��������;2-�쳣�ؽ�;3-�쳣����,4-�������ݲ鿴;5-���ϵ��ݲ鿴;6-�쳣��������
    '����:objBalanceItems_Out-����������
    '     objBalanceInfor-��ǰ������Ϣ
    '     strErrMsg_Out-���ش�����Ϣ
    '     blnҽ����������_Out-ҽ�������Ƿ�ȫ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-03-30 10:31:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPtItems As clsBalanceItems, objItems As clsBalanceItems, objItemsTemp As clsBalanceItems
    Dim objMulitItems As clsBalanceItems, objTransItems As clsBalanceItems
    Dim objItemTemp As clsBalanceItem, objItem As clsBalanceItem
    Dim blnת�� As Boolean, blnSingleDel As Boolean, strErrMsg As String, strExpend As String
    Dim cllDelSwap As Collection, blnNoBalanceData As Boolean '�Ƿ��в���Ԥ����¼
    Dim strTemp As String, strDefaultBalance As String
    Dim dblMoney As Double, lng�����ID As Long
    Dim blnAdd As Boolean, blnDelCash As Boolean, blnFind As Boolean
    Dim rsThirdDel As ADODB.Recordset, i As Long
    Dim objCard As Card, bln�Ƿ񱣴� As Boolean
    Dim intSign As Integer  '������
    Dim strCardTypes As String, rsThirdDelClone As ADODB.Recordset
    Dim strMulitCardTypeIds As String '��ʽ��׺ϲ��Ŀ����IDs:�����ID,...
    Dim strSingleCardTypeIds As String  '�ֽ��׵Ŀ����IDs:�����ID|Ԥ��ID,...
    Dim blnReturnCash As Boolean
    
    On Error GoTo errHandle
    
    blnҽ����������_Out = True
    
    If Not (bytCurType = 4 Or bytCurType = 5) Then '��Ϊ�鿴ʱ�Ŵ���
        If zlGetThirdDelRecordFromBalanceID(objBalanceInfor.����ID, rsThirdDel) = False Then Exit Function
    End If
    
    If objBalanceInfor Is Nothing Then Set objBalanceInfor = New clsBalanceInfo
    Set objBalanceItems_out = New clsBalanceItems
    
    objBalanceInfor.��ǰ���� = 0
    objBalanceInfor.�Ѹ��ϼ� = 0
    objBalanceInfor.ҽ��֧���ϼ� = 0
    
    
    strTemp = "  δ�ҵ�ԭʼ�Ľ����¼"
    If rsBalanceRecord Is Nothing Then strErrMsg_out = strTemp: Exit Function
    If rsBalanceRecord.State <> 1 Then strErrMsg_out = strTemp: Exit Function
    
    bln�Ƿ񱣴� = IIf(bytCurType = 1 Or bytCurType = 4 Or bytCurType = 5, False, True) '����ʱ����δ���г�������.
    
    Set objPtItems = New clsBalanceItems
    
    rsBalanceRecord.Sort = "����,�����ID,��������ID"
    intSign = IIf(bytCurType = 3 Or bytCurType = 5, -1, 1)
    
    
     Set objItems = Nothing
    'objBalanceInfor.��Ԥ���ϼ� = 0
    With rsBalanceRecord
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            objBalanceInfor.��ǰ���� = objBalanceInfor.��ǰ���� + intSign * Val(NVL(!��Ԥ��))
            
            lng�����ID = Val(NVL(!�����ID))
            Select Case Val(NVL(!����))
            Case 0  '��ͨ����
                If NVL(!���㷽ʽ) <> "" Then
                    Set objItem = New clsBalanceItem
                    Set objCard = zlGetCardFromCardType(0, False, NVL(!���㷽ʽ))
                    If Not zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) Then Exit Function
                    objItem.����ID = objBalanceInfor.����ID
                    objItem.����IDs = objBalanceInfor.����ID
                    objItem.����ID = objBalanceInfor.����ID
                    objItem.����ʱ�� = objBalanceInfor.����ʱ��
                    objItem.�������� = Val(NVL(!����))
                    objItem.������ = RoundEx(intSign * objItem.������, 6)
                     
                    objItem.�Ƿ�����ɾ�� = True
                    objItem.�Ƿ��������� = True
                    objItem.�Ƿ�����༭ = False
                    
                    
                    If Val(NVL(!����)) = 1 And (bytCurType <> 4 And bytCurType <> 5) Then '�ֽ����⴦��
                        objBalanceInfor.�ֽ�֧�� = Val(NVL(!��Ԥ��))
                    Else
                        objBalanceItems_out.AddItem objItem
                        objBalanceInfor.�Ѹ��ϼ� = RoundEx(objBalanceInfor.�Ѹ��ϼ� + objItem.������, 6)
                        objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
                    End If
                End If
            Case 1 'Ԥ����
                'objBalanceInfor.��Ԥ���ϼ� = objBalanceInfor.��Ԥ���ϼ� + Val(nvl(!��Ԥ��))
                'objBalanceInfor.�Ƿ񱣴�Ԥ�� = True
            Case 2 'ҽ��
                Set objCard = zlGetCardFromCardType(lng�����ID, Val(NVL(!����)) = 5, NVL(!���㷽ʽ))
                If zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) = False Then Exit Function
                blnAdd = True
                If bln���� Then
                       Select Case Val(NVL(!����))
                       Case 3   '�����ʻ�
                            If bytMCMode = 1 And Not objThirdSwap.zlGetYbPara.���ﲡ�˽������� Then
                                '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                                blnAdd = False '����
                            Else
                                blnAdd = gclsInsure.GetCapability(IIf(bytMCMode = 1, support�����������, supportסԺ��������), lng����ID, intInsure, NVL(!���㷽ʽ))
                            End If
                       Case 4  'ҽ������
                            If bytMCMode = 1 And Not objThirdSwap.zlGetYbPara.���ﲡ�˽������� Then
                                blnAdd = True 'ԭ���˻�
                            Else
                                blnAdd = gclsInsure.GetCapability(IIf(bytMCMode = 1, support�����������, supportסԺ��������), lng����ID, intInsure, NVL(!���㷽ʽ))  '����Ƿ�֧��ԭ����
                            End If
                       End Select
                End If
                
                If blnAdd = False Then blnҽ����������_Out = False
                
                If blnAdd Then
                    objItem.�������� = Val(NVL(!����))
                    objItem.����ID = objBalanceInfor.����ID
                    objItem.����IDs = objBalanceInfor.����ID
                    objItem.����ʱ�� = objBalanceInfor.����ʱ��
                    objItem.����ID = objBalanceInfor.����ID
                    objItem.�Ƿ�����ɾ�� = False
                    objItem.�Ƿ��������� = False
                    objItem.�Ƿ�����༭ = False
                    objItem.�Ƿ񱣴� = bln�Ƿ񱣴�
                    objItem.������ = RoundEx(intSign * objItem.������, 6)
                    objBalanceItems_out.AddItem objItem
                    
                    objBalanceInfor.�Ѹ��ϼ� = RoundEx(objBalanceInfor.�Ѹ��ϼ� + objItem.������, 6)
                    objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 5)
                    If objItem.У�Ա�־ <> 1 Then
                        objBalanceInfor.ҽ��֧���ϼ� = RoundEx(objBalanceInfor.ҽ��֧���ϼ� + objItem.������, 5)
                    End If
                End If
            Case 3 'һ��ͨ
                
                If lng�����ID = 0 Then strErrMsg_out = "��������������������(" & NVL(!���㷽ʽ) & ")�����IDΪ��(,����ϵͳ����Ա��ϵ!": Exit Function
                Set objCard = zlGetCardFromCardType(lng�����ID, False, NVL(!���㷽ʽ))
                If zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) = False Then Exit Function
                
                objItem.�������� = Val(NVL(!����))
                objItem.����IDs = objBalanceInfor.����ID
                objItem.����ID = objBalanceInfor.����ID
                objItem.����ID = objBalanceInfor.����ID
                objItem.����ʱ�� = objBalanceInfor.����ʱ��
                objItem.������ = RoundEx(intSign * objItem.������, 6)
                objItem.�Ƿ�����ɾ�� = Val(NVL(!У�Ա�־)) = 1
                objItem.�Ƿ��������� = False
                objItem.�Ƿ�����༭ = False
                objItem.Ԥ��ID = Val(NVL(!Ԥ��ID))
                objItem.�Ƿ񱣴� = bln�Ƿ񱣴�
                
                blnAdd = True
                Select Case bytCurType
                Case 0, 1, 2, 4, 6 '0-����;1-��������;2-�쳣�ؽ�;3-�쳣����,4-�������ݲ鿴;5-���ϵ��ݲ鿴;6-�쳣��������
                     
                    If objItem.������ < 0 And objBalanceInfor.����ID = 0 Then '������Ϣ
                        
                        objItem.�Ƿ�Ԥ�� = True: objItem.�Ƿ��˿� = True
                        '0-��ͨҵ��;1-�ֽ����˿�,2-����һ�ν��׽ӿ��˿�;3-ת�ʷ�ʽ�˿�
                        Select Case Val(NVL(!���ӱ�־))
                        Case 2 '����һ�ν��׽ӿ��˿�
                            If InStr(strMulitCardTypeIds & ",", "," & lng�����ID & ",") = 0 Then strMulitCardTypeIds = strMulitCardTypeIds & "," & lng�����ID
                            blnFind = False
                            For i = 1 To objBalanceItems_out.Count
                                 If objItem.�����ID = objBalanceItems_out(i).�����ID Then
                                     If objBalanceItems_out(i).objTag Is Nothing Then
                                         Set objBalanceItems_out(i).objTag = New clsBalanceItems
                                     End If
                                     Set objItems = objBalanceItems_out(i).objTag
                                     objItems.AddItem objItem
                                     objItems.������ = RoundEx(objItems.������ + objItem.������, 6)
                                     objBalanceItems_out(i).������ = RoundEx(objBalanceItems_out(i).������ + objItem.������, 6)
                                     objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
                                     blnFind = True
                                     Exit For
                                 End If
                            Next
                            If Not blnFind Then
                                 Set objItemTemp = zlCopyNewItemFromBalanceItem(objItem)
                                 Set objItemTemp.objTag = New clsBalanceItems
                                 Set objItems = objItemTemp.objTag
                                 objItemTemp.�Ƿ��˿�ֽ��� = False
                                 objItems.������ = RoundEx(objItems.������ + objItem.������, 6)
                                 objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
                                 objItems.AddItem objItem
                                 objBalanceItems_out.AddItem objItemTemp
                            End If
                            blnAdd = False
                        Case 1   '1-�ֽ����˿�
                            strTemp = lng�����ID & "|" & objItem.Ԥ��ID
                            If InStr(strSingleCardTypeIds & ",", "," & strTemp & ",") = 0 Then strSingleCardTypeIds = strSingleCardTypeIds & "," & strTemp
                            If Not (bytCurType = 4 Or bytCurType = 5) Then  '�����ڲ鿴����Ҫ����ת�˻�ֽ����˿�
                                If zlGetDelBalanceItemsFromRecord(objItem, rsThirdDel, objItemsTemp, True) = False Then Exit Function
                                objItem.�Ƿ��˿�ֽ��� = True
                                If objItem.Ԥ��ID <> objItemsTemp(1).Ԥ��ID Then objItem.Ԥ��ID = objItemsTemp(1).Ԥ��ID
                                Set objItem.objTag = objItemsTemp
                            End If
                        Case Else   '0-��ͨҵ��;1-�ֽ����˿�,2-����һ�ν��׽ӿ��˿�;3-ת�ʷ�ʽ�˿�
                            If InStr(strMulitCardTypeIds & ",", "," & lng�����ID & ",") = 0 Then strMulitCardTypeIds = strMulitCardTypeIds & "," & lng�����ID
                            If Not (bytCurType = 4 Or bytCurType = 5) Then  '�����ڲ鿴����Ҫ����ת�˻�ֽ����˿�
                                If zlGetDelBalanceItemsFromRecord(objItem, rsThirdDel, objItemsTemp) = False Then Exit Function
                                Set objItem.objTag = objItemsTemp
                            End If
                        End Select
                    ElseIf objItem.������ > 0 And bytCurType = 1 Then '��������ʱ����Ҫ�ȼ���Ƿ�ȱʡ
                            
                        If InStr(strMulitCardTypeIds & ",", "," & lng�����ID & ",") = 0 Then strMulitCardTypeIds = strMulitCardTypeIds & "," & lng�����ID
                        
                        strCardTypes = !���� & "_" & Val(NVL(!�����ID)) & "_" & Val(NVL(!��������ID))
                        Set objItemsTemp = New clsBalanceItems
                        i = 0
                        objItem.�Ƿ��˿� = True
                        Do While Not .EOF
                            If strCardTypes <> !���� & "_" & Val(NVL(!�����ID)) & "_" & Val(NVL(!��������ID)) Then .MovePrevious: Exit Do
                            If i <> 0 Then
                                Set objItem = New clsBalanceItem
                                If zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) = False Then Exit Function
                                objItem.�������� = Val(NVL(!����))
                                objItem.����IDs = objBalanceInfor.����ID
                                objItem.����ID = objBalanceInfor.����ID
                                objItem.����ID = objBalanceInfor.����ID
                                objItem.����ʱ�� = objBalanceInfor.����ʱ��
                                objItem.�Ƿ�����ɾ�� = False
                                objItem.�Ƿ��������� = False
                                objItem.�Ƿ�����༭ = False
                                objItem.�Ƿ��˿� = True
                                objItem.�Ƿ񱣴� = bln�Ƿ񱣴�
                                objBalanceInfor.��ǰ���� = RoundEx(objBalanceInfor.��ǰ���� + objItem.������, 6)
                            End If
                            i = i + 1
                            objItemsTemp.AddItem objItem
                            objItemsTemp.������ = objItemsTemp.������ + objItem.������
                            .MoveNext
                        Loop
                        If .EOF Then .MovePrevious
                        
                        For i = 1 To objItemsTemp.Count
                            '���д����,���⴫��ӿ�Ϊ����
                            objItemsTemp(i).������ = RoundEx(-1 * objItemsTemp(i).������, 6)
                        Next
                        '���ò���"�����ʽ������µ�Ԥ����"ʱ,���������ֽӿ�
                        If Not (bytCurType = 1 And gTy_System_Para.TY_Balance.bln�������ϲ�����Ԥ�� And gTy_System_Para.TY_Balance.str����Ԥ�����㷽ʽ <> "") Then
                            blnReturnCash = objThirdSwap.zlThirdReturnCashCheck(objCard, objItemsTemp, blnDelCash, strDefaultBalance)
                        End If
                        If Not blnReturnCash Then
                            For i = 1 To objItemsTemp.Count
                                objItemsTemp(i).�Ƿ��������� = False
                                objItemsTemp(i).�Ƿ�ǿ������ = blnDelCash
                                objItemsTemp(i).�Ƿ�����ɾ�� = objItemsTemp(i).�Ƿ�ǿ������
                                objItemsTemp(i).������ = RoundEx(-1 * objItemsTemp(i).������, 6)
                                objItemsTemp(i).�Ƿ��˿� = True
                                objItemsTemp(i).����ID = objBalanceInfor.����ID
                                objItemsTemp(i).����ID = objBalanceInfor.����ID
                                objItemsTemp(i).����ʱ�� = objBalanceInfor.����ʱ��
                                objItemsTemp(i).�Ƿ񱣴� = bln�Ƿ񱣴�
                                objBalanceItems_out.AddItem objItemsTemp(i)
                                objBalanceInfor.�Ѹ��ϼ� = RoundEx(objBalanceInfor.�Ѹ��ϼ� + objItemsTemp(i).������, 6)
                                objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItemsTemp(i).������, 6)
                            Next
                            blnAdd = False
                        Else
                            For i = 1 To objItemsTemp.Count
                                '���д����,���⴫��ӿ�Ϊ����
                                objItemsTemp(i).������ = RoundEx(-1 * objItemsTemp(i).������, 6)
                            Next
                        
                            blnAdd = False
                            If blnDelCash = False Then  '�Ƿ�ȱʡ����
                                For i = 1 To objItemsTemp.Count
                                    objItemsTemp(i).�Ƿ��������� = True
                                    objItemsTemp(i).�Ƿ�����ɾ�� = True
                                    objItemsTemp(i).�Ƿ�ǿ������ = True
                                    objItemsTemp(i).����ID = objBalanceInfor.����ID
                                    objItemsTemp(i).����ID = objBalanceInfor.����ID
                                    objItemsTemp(i).����ʱ�� = objBalanceInfor.����ʱ��
                                    objItemsTemp(i).�Ƿ񱣴� = bln�Ƿ񱣴�
                                    
                                    objBalanceItems_out.AddItem objItemsTemp(i)
                                    objBalanceInfor.�Ѹ��ϼ� = RoundEx(objBalanceInfor.�Ѹ��ϼ� + objItemsTemp(i).������, 6)
                                    objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItemsTemp(i).������, 6)
                                Next
                           ElseIf strDefaultBalance <> "" Then
                                Set objItemTemp = New clsBalanceItem
                                With objItemTemp
                                    Set .objCard = zlGetCardFromBalanceName(strDefaultBalance)
                                    .���㷽ʽ = strDefaultBalance
                                    .������ = RoundEx(intSign * objItemsTemp.������, 6)
                                    .�Ƿ��˿� = True
                                    .�Ƿ�����༭ = False
                                    .�Ƿ�����ɾ�� = True
                                    .�������� = .objCard.��������
                                    .����IDs = objBalanceInfor.����ID
                                    .����ID = objBalanceInfor.����ID
                                    .����ID = objBalanceInfor.����ID
                                    .����ʱ�� = objBalanceInfor.����ʱ��
                                    .�Ƿ񱣴� = bln�Ƿ񱣴�
                                End With
                                objPtItems.AddItem objItemTemp
                                objPtItems.������ = RoundEx(objPtItems.������ + objItem.������, 6)
                            End If
                         
                        End If
                    End If
                End Select
                
                If blnAdd Then
                    objBalanceItems_out.AddItem objItem
                    objBalanceInfor.�Ѹ��ϼ� = RoundEx(objBalanceInfor.�Ѹ��ϼ� + objItem.������, 6)
                    objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
                End If
                
            Case 4 '��һ��ͨ
                
                Set objCard = zlGetCardFromCardType(lng�����ID, False, NVL(!���㷽ʽ))
                If zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) = False Then Exit Function
                
                objItem.�������� = Val(NVL(!����))
                objItem.����IDs = objBalanceInfor.����ID
                objItem.����ID = objBalanceInfor.����ID
                objItem.����ID = objBalanceInfor.����ID
                objItem.����ʱ�� = objBalanceInfor.����ʱ��
                objItem.�Ƿ񱣴� = bln�Ƿ񱣴�
                objItem.������ = RoundEx(intSign * objItem.������, 6)
                If objItem.������ < 0 And objBalanceInfor.����ID = 0 Then  '�����˿�
                    '��ȡ��ϸ��
                    objItem.�Ƿ�Ԥ�� = True: objItem.�Ƿ��˿� = True
                End If
                
                objItem.�Ƿ񱣴� = bln�Ƿ񱣴�
                objBalanceItems_out.AddItem objItem
                objBalanceInfor.�Ѹ��ϼ� = RoundEx(objBalanceInfor.�Ѹ��ϼ� + objItem.������, 6)
                objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
            Case 5 '���ѿ�
                lng�����ID = Val(NVL(!���㿨���))
                Set objCard = zlGetCardFromCardType(lng�����ID, True, NVL(!���㷽ʽ))
                If zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) = False Then Exit Function
                objItem.�����ID = lng�����ID
                objItem.�������� = Val(NVL(!����))
                objItem.����ID = objBalanceInfor.����ID
                objItem.����IDs = objBalanceInfor.����ID
                objItem.����ID = objBalanceInfor.����ID
                Select Case bytCurType
                Case 0, 1, 2, 4, 6 '0-����;1-��������;2-�쳣�ؽ�;3-�쳣����,4-�������ݲ鿴;5-���ϵ��ݲ鿴;6-�쳣��������
                    If objItem.������ < 0 And objBalanceInfor.����ID = 0 Then objItem.�Ƿ�Ԥ�� = True
                End Select
                objItem.���ѿ� = True
                objItem.����ʱ�� = objBalanceInfor.����ʱ��
                objItem.�Ƿ񱣴� = bln�Ƿ񱣴�
                objItem.������ = RoundEx(intSign * objItem.������, 6)
                objBalanceItems_out.AddItem objItem
                objBalanceInfor.�Ѹ��ϼ� = RoundEx(objBalanceInfor.�Ѹ��ϼ� + objItem.������, 6)
                objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
            Case Else
                '���������ѣ�������
            End Select
            rsBalanceRecord.MoveNext
        Loop
    End With
    
    '������Ϣ�п��ܲ��������������˿���Ϣ�У�Ĭ��Ӧ�ö�ȡ����
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    Set cllDelSwap = New Collection

    If Not (bytCurType = 4 Or bytCurType = 5 Or bytCurType = 6) Then '��Ϊ�鿴���쳣����ʱ�����ü����ⲿ����������

        rsThirdDel.Filter = ""
        Set rsThirdDelClone = zlDatabase.CopyNewRec(rsThirdDel)

        With rsThirdDel
            .Filter = 0
            Do While Not .EOF

                lng�����ID = Val(NVL(!�����ID))
                blnת�� = Val(NVL(rsThirdDel!�Ƿ�ת��)) = 1
                strTemp = lng�����ID & "|" & Val(NVL(!Ԥ��ID))

                If Not (InStr(1, strMulitCardTypeIds & ",", "," & lng�����ID & ",") > 0 Or InStr(strSingleCardTypeIds & ",", "," & strTemp & ",") > 0) Then

                    Set objCard = zlGetCardFromCardType(lng�����ID, False, "")
                    Set objItem = New clsBalanceItem
                    With objItem
                        .��������ID = 0
                        .������ = 0
                        .���㷽ʽ = objCard.���㷽ʽ
                        .�������� = 3

                        .����IDs = objBalanceInfor.����ID
                        .����ID = objBalanceInfor.����ID
                        .����ID = objBalanceInfor.����ID
                        .�������� = objCard.��������
                        .����ʱ�� = objBalanceInfor.����ʱ��
                        .�����ID = lng�����ID
                        .�Ƿ񱣴� = bln�Ƿ񱣴�
                        .�Ƿ��˿� = True
                        .�Ƿ���� = Val(NVL(rsThirdDel!�Ƿ�δ��)) = 0
                        .�Ƿ�Ԥ�� = True
                        .�Ƿ����� = objCard.�������Ĺ��� <> ""
                        .У�Ա�־ = IIf(Val(NVL(rsThirdDel!�Ƿ�δ��)) = 1, 1, 2)
                        .�Ƿ�ת�� = blnת��
                        .Ԥ��ID = Val(NVL(rsThirdDel!Ԥ��ID))
                        .�Ƿ�����ɾ�� = Val(NVL(rsThirdDel!�Ƿ�δ��)) = 1

                        Set .objCard = objCard
                    End With
                    
                    blnNoBalanceData = True
                    If Not (blnת�� Or InStr(1, strSingleCardTypeIds & "|", "," & lng�����ID & "|") > 0) Then
                        blnFind = False
                        For i = 1 To cllDelSwap.Count
                           If cllDelSwap(i)(0) = lng�����ID Then
                                blnFind = True
                                blnSingleDel = cllDelSwap(i)(1) = 0: Exit For
                           End If
                        Next

                        If Not blnFind And (bytCurType = 2 Or bytCurType = 3) Then
                            blnSingleDel = ThirdSwapIsSwapNOCall(objBalanceInfor.����ID, lng�����ID, blnNoBalanceData)
                            If blnNoBalanceData Then
                                '��Ҫ����ȷ���Ƿ�ֽ��׵���
                                blnSingleDel = objThirdSwap.zlThirdSwapIsSwapNOCall(lng�����ID, False, strErrMsg, strExpend)
                            End If
                            cllDelSwap.Add Array(lng�����ID, IIf(blnSingleDel, 0, 1))
                        End If
                        objItem.�Ƿ��˿�ֽ��� = blnSingleDel
                    ElseIf InStr(1, strSingleCardTypeIds & "|", "," & lng�����ID & "|") > 0 Then
                        objItem.�Ƿ��˿�ֽ��� = True
                    End If

                    If objItem.�Ƿ��˿�ֽ��� Then
                        objItem.��������ID = Val(NVL(rsThirdDel!��������ID))
                        objItem.������ˮ�� = NVL(rsThirdDel!������ˮ��)
                        objItem.����˵�� = NVL(rsThirdDel!����˵��)
                    End If

                    If zlGetDelBalanceItemsFromRecord(objItem, rsThirdDelClone, objItemsTemp, objItem.�Ƿ��˿�ֽ���) = False Then Exit Function

                    Set objItem.objTag = objItemsTemp
                    objItem.������ = objItemsTemp.������

                    'û�в���Ԥ����¼�Ķ�����
                    If (bytCurType = 2 Or bytCurType = 3) And blnNoBalanceData Then
                        '���ò���"�����ʽ������µ�Ԥ����"ʱ,���������ֽӿ�
                        If Not (bytCurType = 3 And gTy_System_Para.TY_Balance.bln�������ϲ�����Ԥ�� And gTy_System_Para.TY_Balance.str����Ԥ�����㷽ʽ <> "") Then
                            blnReturnCash = objThirdSwap.zlThirdReturnCashCheck(objCard, objItemsTemp, blnDelCash, strDefaultBalance)
                        End If
                        If Not blnReturnCash Then
                            objItem.�Ƿ��������� = False
                            objItem.�Ƿ�ǿ������ = blnDelCash
                            objItem.�Ƿ�����ɾ�� = objItem.�Ƿ�ǿ������
                            blnAdd = True
                        Else
                            objItem.�Ƿ��������� = True: objItem.�Ƿ�ǿ������ = True
                            objItem.�Ƿ�����ɾ�� = True
                            If blnDelCash = False Then  '�Ƿ�ȱʡ����
                                objItem.�Ƿ�����༭ = False
                                objItem.�Ƿ�����ɾ�� = True
                                blnAdd = True
                            ElseIf strDefaultBalance <> "" Then
                                Set objItemTemp = New clsBalanceItem
                                With objItemTemp
                                    Set .objCard = zlGetCardFromBalanceName(strDefaultBalance)
                                    .���㷽ʽ = strDefaultBalance
                                    .������ = RoundEx(objItem.������, 6)
                                    .�Ƿ��˿� = True
                                    .�Ƿ�����༭ = False
                                    .�Ƿ�����ɾ�� = True
                                    .�������� = .objCard.��������
                                    .����IDs = objBalanceInfor.����ID
                                    .����ID = objBalanceInfor.����ID
                                    .����ID = objBalanceInfor.����ID
                                    .����ʱ�� = objBalanceInfor.����ʱ��
                                    .�Ƿ��������� = True
                                    .�Ƿ�ǿ������ = True
                                End With
                                objPtItems.AddItem objItemTemp
                                objPtItems.������ = RoundEx(objPtItems.������ + objItemTemp.������, 6)
                                blnAdd = False
                            End If
                        End If
                    Else
                        blnAdd = True
                    End If

                    If blnAdd Then
                        objBalanceItems_out.AddItem objItem
                    End If

                    If objItem.�Ƿ��˿�ֽ��� Then
                        strSingleCardTypeIds = strSingleCardTypeIds & "," & strTemp
                    Else
                        strMulitCardTypeIds = strMulitCardTypeIds & "," & lng�����ID
                    End If
                End If
                .MoveNext
            Loop
        End With
    End If
    
    '������ͨ�Ľ��㷽ʽ
    For Each objItem In objPtItems
        blnAdd = True
        For Each objItemTemp In objBalanceItems_out
            If objItemTemp.���㷽ʽ = objItem.���㷽ʽ And objItemTemp.�������� = 0 Then
                objItemTemp.������ = RoundEx(objItemTemp.������ + objItem.������, 6)
                objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
                blnAdd = False
                Exit For
            End If
        Next
        If blnAdd Then
            objBalanceItems_out.AddItem objItem
            objBalanceItems_out.������ = RoundEx(objBalanceItems_out.������ + objItem.������, 6)
        End If
    Next
    objBalanceInfor.δ���ϼ� = RoundEx(objBalanceInfor.��ǰ���� - objBalanceInfor.�Ѹ��ϼ�, 6)
    objBalanceInfor.�Ƿ񱣴���ʵ� = bln�Ƿ񱣴�
    
    zlGetBalanceItemsFromRecord = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ThirdSwapIsSwapNOCall(ByVal lng����ID As Long, ByVal lng�����ID As Long, ByRef blnNoData As Boolean) As Boolean
    '�ж��Ƿ�ֵ��ݽ���
    '���:
    '����:
    '   blnNoData-�Ƿ��޽�������
    '˵��:
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    blnNoData = False
    strSQL = _
        " Select ���ӱ�־ From ����Ԥ����¼" & _
        " Where ��¼���� = 2 And ����id = [1] And �����id = [2] And ��Ԥ�� < 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ�ֵ��ݽ���", lng����ID, lng�����ID)
    If rsTemp.EOF Then blnNoData = True: Exit Function
    ThirdSwapIsSwapNOCall = Val(NVL(rsTemp!���ӱ�־)) = 1 '�ֽ����˿�
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetBalanceItemFromRecord(ByVal objCard As Card, ByVal rsBalanceRecord As ADODB.Recordset, _
    ByRef objBalanceItem_Out As clsBalanceItem, Optional strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ��ʼ�¼���ݷ���ָ���Ľ�����Ϣ��
    '���:objCard-������
    '     int����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    '     lng��������Id-��������ID
    '����:objBalanceItem_Out-���ؽ�����Ϣ��
    '     strErrMsg_Out-���ش�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-03-30 10:31:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Dim dblMoney As Double
    
    Set objBalanceItem_Out = New clsBalanceItem
    
    On Error GoTo errHandle
    
    strTemp = "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!"
    If rsBalanceRecord Is Nothing Then strErrMsg_out = strTemp: Exit Function
    If rsBalanceRecord.State <> 1 Then strErrMsg_out = strTemp: Exit Function
    If rsBalanceRecord.EOF Then Exit Function
    With rsBalanceRecord
        With objBalanceItem_Out
            Set .objCard = objCard
            .���㷽ʽ = NVL(rsBalanceRecord!���㷽ʽ)
            .������ = Val(NVL(rsBalanceRecord!��Ԥ��))
            .��������ID = Val(NVL(rsBalanceRecord!��������ID))
            .������ˮ�� = NVL(rsBalanceRecord!������ˮ��)
            .����˵�� = NVL(rsBalanceRecord!����˵��)
            .������� = NVL(rsBalanceRecord!�������)
            .�������� = Val(NVL(rsBalanceRecord!����))
            .����ժҪ = NVL(rsBalanceRecord!ժҪ)
            .���� = NVL(rsBalanceRecord!����)
            .�����ID = Val(NVL(rsBalanceRecord!�����ID))
            .���ѿ�ID = Val(NVL(rsBalanceRecord!���ѿ�ID))
            .���ѿ� = Val(NVL(rsBalanceRecord!���ѿ�ID)) <> 0
            .�Ƿ����� = Val(NVL(rsBalanceRecord!�Ƿ�����)) = 1
            .ԭʼ��� = Val(NVL(rsBalanceRecord!��Ԥ��))
            .δ�˽�� = Val(NVL(rsBalanceRecord!��Ԥ��))
            .�Ƿ��˿� = Val(NVL(rsBalanceRecord!��Ԥ��)) < 0
            .�Ƿ�����༭ = False
            .�Ƿ�����ɾ�� = False
            .У�Ա�־ = Val(NVL(rsBalanceRecord!У�Ա�־))
            .�Ƿ���� = Val(NVL(rsBalanceRecord!У�Ա�־)) = 2 Or Val(NVL(rsBalanceRecord!У�Ա�־)) = 0
            '.����ʱ�� =  Format(rsBalanceRecord!�տ�ʱ��, "yyyy-mm-dd HH:MM:SS")
            .������� = "" '  Nvl(rsBalanceRecord!�������)
            .���� = ""
            .�ʻ���� = 0
            .�������� = Val(NVL(rsBalanceRecord!����))
            .����IDs = 0 'Val(nvl(rsBalanceRecord!����ID))
        End With
    End With
    zlGetBalanceItemFromRecord = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlInsureCheck(ByVal str���ս��� As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ��ҽ���Ƿ���Ҫ�϶�
    '���:str���ս���-���ս���
    '       strAdvance-ҽ�����صĽ���
    '����:
    '����:��Ҫ�϶�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo errHandle
    If Not (strAdvance <> "" And str���ս��� <> strAdvance) Then Exit Function
    '��ʽ����ǰ��,���㷽ʽ�ͽ�����δ�����仯ʱ��У��
    blnMedicareCheck = True
    varData = Split(str���ս���, "||"): varData1 = Split(strAdvance, "||")

    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsureCheck = blnMedicareCheck
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckInsureCancelIsValied(ByVal lng����ID As Long, ByVal str���Ͻ��㷽ʽs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���Ľ�����㷽ʽ����Ƿ�ԭ����
    '���:lng����ID-����ID
    '     str���㷽ʽs-�������ϵĽ�����Ϣ,����ö��ŷ���
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-17 13:45:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str������Ϣ As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select  ���㷽ʽ From ����Ԥ����¼ A,���㷽ʽ B  " & vbCrLf & _
    "   Where a.����ID=[1] and a.���㷽ʽ=B.���� and b.���� in (3,4)  and mod( A.��¼����,10)<>1  " & vbCrLf & _
    "         And nvl(a.�����ID,0)=0 And ���㷽ʽ not in (Select Column_value From table(f_str2List([2])))" & vbCrLf
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ���Ƿ����ԭʼ������Ϣ", lng����ID, str���Ͻ��㷽ʽs)
    If rsTemp.EOF Then zlCheckInsureCancelIsValied = True: Exit Function
    
    str������Ϣ = ""
    With rsTemp
        Do While Not .EOF
            str������Ϣ = str������Ϣ & vbCrLf & NVL(!���㷽ʽ)
            .MoveNext
        Loop
    End With
    MsgBox "��ҽ����֧��ԭ���˻ش���,���������ϣ���֧�����ϵĽ�����Ϣ����:" & str������Ϣ, vbInformation + vbOKOnly, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If


End Function
Public Function zlGetBalanceItemFromCardObject(ByVal objCurCard As Card, ByVal dblMoney As Double, ByRef objItem_Out As clsBalanceItem, _
    Optional str����ժҪ As String, Optional str������� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ����󣬻�ȡ�µĽ�����Ϣ����
    '���:objCurCard-��ǰ������
    '     dblMoney-��ǰ������
    '����:objItem_Out-��ǰ��������
    '     objCurCard-���ص�ǰ��֧�������
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-23 19:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, int����    As Integer
    On Error GoTo errHandle
    
    If objCurCard Is Nothing Then Exit Function
    
    int���� = IIf(objCurCard.�ӿ���� > 0, IIf(objCurCard.���ѿ�, 5, 3), 0)
    If objCurCard.�������� = 7 Then int���� = 4
    
    Set objItem_Out = New clsBalanceItem
    With objItem_Out
        Set .objCard = objCurCard: Set .objTag = Nothing
        .��������ID = 0
        .����IDs = ""
        .���㷽ʽ = objCurCard.���㷽ʽ
        .������� = str�������
        .������ = dblMoney
        .�������� = int����  '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        .�������� = objCurCard.��������
        .����ժҪ = str����ժҪ
        .���� = ""
        .�����ID = IIf(objCurCard.�ӿ���� > 0, objCurCard.�ӿ����, 0)
        .�Ƿ����� = objCurCard.�������Ĺ��� <> ""
        .�Ƿ�����༭ = False
        .�Ƿ�����ɾ�� = False
        .�Ƿ��������� = False
        .�Ƿ�ת�� = False
        .������� = ""
        .���ѿ� = objCurCard.���ѿ�
        .У�Ա�־ = 1
    End With
    zlGetBalanceItemFromCardObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceItemsFromVsBalanceGrid(ByVal vsBalance As VSFlexGrid, ByVal objCurItem As clsBalanceItem, ByRef objItems_Out As clsBalanceItems, _
    Optional ByVal blnNewItem As Boolean = False, Optional blnSign As Boolean, Optional str���㷽ʽs_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ������񣬻�ȡָ�����ݼ�
    '���:vsBalance-�����б�
    '     objItem-��ǰ������Ϣ����
    '     blnNewItem-�Ƿ񷵻ص�ֵ��ȫ�µ�ֵ
    '     blnSign-����Ƿ�ȡ�෴��,true-ȡ�෵��,����ԭʼֵ
    '����:objItems_Out-���صĹ�����
    '      str���㷽ʽs_out:���ؽ�����Ϣ:��ʽ:���㷽ʽ:������|���㷽ʽ:������;
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-13 16:56:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As clsBalanceItem, objNewItem As clsBalanceItem
    Dim intSign As Integer, str���㷽ʽs As String
    On Error GoTo errHandle
    
    intSign = IIf(blnSign, -1, 1)
    str���㷽ʽs_out = ""
    Set objItems_Out = New clsBalanceItems
    If objCurItem.��������ID = 0 Then  '�޹�������ID,��ֱ�ӷ���s
        If blnNewItem Then
            Set objNewItem = zlCopyNewItemFromBalanceItem(objCurItem)
        Else
            Set objNewItem = objCurItem
        End If
        objNewItem.������ = RoundEx(intSign * objNewItem.������, 6)
        str���㷽ʽs_out = objNewItem.���㷽ʽ & ":" & Format(objNewItem.������, "0.00")
        objItems_Out.AddItem objNewItem
        objItems_Out.������ = objNewItem.������
        zlGetBalanceItemsFromVsBalanceGrid = True
        
        Exit Function
    End If
    With vsBalance
        For i = 1 To .Rows - 1
             If zlGetBalanceItemFromBalanceGrid(vsBalance, i, objItem) Then
                If objCurItem.��������ID = objItem.��������ID And objCurItem.�����ID = objItem.�����ID And objCurItem.Ԥ��ID = objItem.Ԥ��ID Then
                    If blnNewItem Then
                        Set objNewItem = zlCopyNewItemFromBalanceItem(objItem)
                    Else
                        Set objNewItem = objItem
                    End If
                    objNewItem.������ = RoundEx(intSign * objNewItem.������, 6)
                    objItems_Out.AddItem objNewItem
                    objItems_Out.�Ƿ�ת�� = objNewItem.�Ƿ�ת��
                    objItems_Out.�շ����� = IIf(objNewItem.�Ƿ�Ԥ��, 1, 0)
                    objItems_Out.�շ����� = objNewItem.��������
                    objItems_Out.������ = RoundEx(objItems_Out.������ + objNewItem.������, 6)
                    str���㷽ʽs_out = str���㷽ʽs_out & "|" & objNewItem.���㷽ʽ & ":" & Format(objNewItem.������, "0.00")
                End If
             End If
        Next
        If str���㷽ʽs_out <> "" Then str���㷽ʽs_out = Mid(str���㷽ʽs_out, 2)
    End With
    
    If objItems_Out.Count = 0 Then Exit Function
    Set objItem = Nothing
    zlGetBalanceItemsFromVsBalanceGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCancelBalancesFromVsBalanceGrid(ByVal vsBalance As VSFlexGrid, ByVal bytFun As Byte, ByRef strBalances As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ϵ���ͨ���㷽ʽ
    '���:bytFun-0-��ͨ;1-ҽ��;2-���ѿ�
    '     vsBalance-�����б�
    '����:
    '    bytfunc=0:strBalances�ĸ�ʽ:���㷽ʽ|������|�������||...
    '    bytfunc=1:strBalances�ĸ�ʽ:���㷽ʽ|������||...
    '    bytfunc=2:strBalances�ĸ�ʽ:�����ID|����|���ѿ�ID|���ѽ��||.
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-22 16:20:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPTBalance As String, i As Long, dblMoney As Double
    Dim strYbBalance As String, strBalance As String, varData As Variant
    Dim strXFBalance As String
    Dim objItem As clsBalanceItem
    
    On Error GoTo errHandle
    With vsBalance
        '�ռ��˿ʽ�����
        strPTBalance = "": strYbBalance = "": strXFBalance = ""
        For i = 1 To .Rows - 1
            dblMoney = -1 * RoundEx(Val(.TextMatrix(i, .ColIndex("������"))), 6)
            strBalance = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
            
            If strBalance <> "" And Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
                '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                Select Case Val(.TextMatrix(i, .ColIndex("����")))
                Case 0 '��ͨ����
                    '���㷽ʽ|������|�������|����ժҪ||..
                    strPTBalance = strPTBalance & "||" & strBalance
                    strPTBalance = strPTBalance & "|" & dblMoney
                    strPTBalance = strPTBalance & "|" & IIf(.TextMatrix(i, .ColIndex("�������")) = "", " ", .TextMatrix(i, .ColIndex("�������")))
                    strPTBalance = strPTBalance & "|" & IIf(.TextMatrix(i, .ColIndex("��ע")) = "", " ", .TextMatrix(i, .ColIndex("��ע")))
                Case 1 'Ԥ����
                Case 2 'ҽ��
                        '���㷽ʽ|������||...
                        strYbBalance = strYbBalance & "||" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & "|" & dblMoney
                Case 3 'һ��ͨ
                Case 4 'һ��ͨ(�ϰ汾)
                Case 5 '���ѿ�
                
                    If zlGetBalanceItemFromBalanceGrid(vsBalance, i, objItem) = False Then Exit Function
                    
                    '�����ID|����|���ѿ�ID|���ѽ��||.
                    strXFBalance = strXFBalance & "||" & objItem.�����ID  ' Val(.TextMatrix(i, .ColIndex("�����ID")))
                    strXFBalance = strXFBalance & "|" & IIf(objItem.���� = "", " ", objItem.����) ' Trim(.Cell(flexcpData, i, .ColIndex("����")))
                    strXFBalance = strXFBalance & "|" & objItem.���ѿ�ID  'Val(.TextMatrix(i, .ColIndex("���ѿ�ID")))
                    strXFBalance = strXFBalance & "|" & dblMoney
                Case Else
                End Select
            End If
        Next
    End With
    If strPTBalance <> "" Then strPTBalance = Mid(strPTBalance, 3)
    If strYbBalance <> "" Then strYbBalance = Mid(strYbBalance, 3)
    If strXFBalance <> "" Then strXFBalance = Mid(strXFBalance, 3)
    
    If bytFun = 0 Then
        strBalances = strPTBalance
    ElseIf bytFun = 1 Then
        strBalances = strYbBalance
    Else
       strBalances = strXFBalance
    End If
    zlGetCancelBalancesFromVsBalanceGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlAddBalanceDataToGridFromBalanceItems(ByVal vsBalance As VSFlexGrid, ByVal objCard As Card, ByRef objBalanceInfor As clsBalanceInfo, ByVal objBalanceItems As clsBalanceItems, Optional lngRow As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ�����Ϣ,���ӽ������ݸ�����
    '���:objCard-��ǰ�Ŀ�����
    '     objBalanceItems-��ǰ�Ľ�����Ϣ��
    '     lngRow-ָ������(���Ϊ0��ʾ�����һ�м���
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-04-10 11:38:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    Dim lngTemp As Long
    
    On Error GoTo errHandle
    
    If objCard.���ѿ� Then Call zlClearSquareBalance(vsBalance, objCard.�ӿ����, objBalanceInfor)        '���ѿ�����Ҫ���ԭ�Ѿ����ڵ�����
    If lngRow <= 0 Then lngRow = zlGetBalanceNULLRow(vsBalance, lngRow)
    If lngRow < 0 Then vsBalance.Rows = vsBalance.Rows + 1: lngRow = vsBalance.Rows - 1
    
    If objBalanceItems Is Nothing Then Exit Sub
    
    With vsBalance
        If .Rows <= 1 Then .Rows = 2
        If lngRow > .Rows - 1 Then
             If Trim(.TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ"))) <> "" Then
                .Rows = .Rows + 1
             End If
             lngRow = .Rows - 1
        End If
        
        For Each objItem In objBalanceItems
              If Trim(.TextMatrix(lngRow, .ColIndex("���㷽ʽ"))) <> "" And Val(.TextMatrix(lngRow, .ColIndex("����״̬"))) <> 0 Then
                '�Ѿ������������,Ӧ�ôӵ�ǰ�в���
                 If lngRow >= .Rows - 1 Then
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                 ElseIf Trim(.TextMatrix(lngRow + 1, .ColIndex("���㷽ʽ"))) <> "" And Val(.TextMatrix(lngRow + 1, .ColIndex("����״̬"))) <> 0 Then
                    '��һ���ǽ����˵����ݣ���Ҫ���м�����У��Ա�ͬһ�ν������һ��
                    lngTemp = zlGetBalanceNULLRow(vsBalance, lngRow)
                    If lngTemp < 0 Then .Rows = .Rows + 1: lngTemp = .Rows = .Rows - 1
                    .RowPosition(lngTemp) = lngRow + 1
                    lngRow = lngRow + 1
                 Else
                    lngRow = lngRow + 1
                 End If
              End If
              
              objItem.�к� = lngRow
              objItem.QRCode = ""
              '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
              .TextMatrix(lngRow, .ColIndex("����")) = objItem.��������
              .TextMatrix(lngRow, .ColIndex("�Ƿ�����")) = IIf(objItem.�Ƿ�����, 1, 0)
              .TextMatrix(lngRow, .ColIndex("��������")) = objCard.��������
              .TextMatrix(lngRow, .ColIndex("�༭״̬")) = IIf(objItem.�Ƿ�����༭, 1, 0) & "|" & IIf(objItem.�Ƿ�����ɾ��, 1, 0)
              .TextMatrix(lngRow, .ColIndex("����״̬")) = IIf(objItem.У�Ա�־ = 2, 1, 0) '�Ƿ��ѽ���:1-�ѽ���;0-δ����
              .TextMatrix(lngRow, .ColIndex("�����ID")) = objItem.�����ID
              .TextMatrix(lngRow, .ColIndex("���ѿ�ID")) = objItem.���ѿ�ID
              .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = objItem.���㷽ʽ
              .TextMatrix(lngRow, .ColIndex("����")) = objCard.zlCardNOEncrypt(objItem.����)
              .TextMatrix(lngRow, .ColIndex("������")) = Format(objItem.������, "0.00")
              .TextMatrix(lngRow, .ColIndex("�������")) = objItem.�������
              .TextMatrix(lngRow, .ColIndex("��ע")) = objItem.����ժҪ
              .TextMatrix(lngRow, .ColIndex("������ˮ��")) = objItem.������ˮ��
              .TextMatrix(lngRow, .ColIndex("����˵��")) = objItem.����˵��
              
              .TextMatrix(lngRow, .ColIndex("�Ƿ�����")) = IIf(objCard.�Ƿ�����, 1, 0)
              .TextMatrix(lngRow, .ColIndex("�Ƿ�ȫ��")) = IIf(objCard.�Ƿ�ȫ��, 1, 0)
              .TextMatrix(lngRow, .ColIndex("���������")) = objCard.����
              
              .Cell(flexcpData, lngRow, .ColIndex("������")) = Format(objItem.������, "0.00")
              .Cell(flexcpData, lngRow, .ColIndex("���ѿ�ID")) = objItem.����
              .Cell(flexcpData, lngRow, .ColIndex("�����ID")) = objItem.�������
              .Cell(flexcpData, lngRow, .ColIndex("����")) = objItem.����
              .Cell(flexcpData, lngRow, .ColIndex("����״̬")) = IIf(objItem.�Ƿ񱣴�, 1, 0)
              
                If objItem.�Ƿ���� Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = g_BalanceRow_Color_Succes
                ElseIf objItem.�Ƿ񱣴� Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = g_BalanceRow_Color_Valied
                Else
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = g_BalanceRow_Color_Normal
                End If
              .RowData(lngRow) = objItem
              If lngRow + 1 > .Rows - 1 Then .Rows = .Rows + 1
              lngRow = lngRow + 1
        Next
        
    End With
    Call zlRecalItemObjectRowNo(vsBalance)    '����ˢ�ж�����к�
   Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlLoadBalanceItemsToVsGrid(ByVal vsGrid As VSFlexGrid, ByVal objBalanceItems As clsBalanceItems, Optional ByVal bln�鿴 As Boolean, Optional ByVal lngRow As Long = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ���������ݼ��ص�����
    '���:lngRow-ָ������:>0ʱ���滻��ǰ��,Ȼ������һ���м�����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-03-29 18:14:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytPreDraw As RedrawSettings
    Dim objBalanceItem As clsBalanceItem
    Dim byt���� As gBalanceType, i As Long
    Dim objCard As Card
    
    On Error GoTo errHandle
    
    If objBalanceItems Is Nothing Then Exit Function '
    
    bytPreDraw = vsGrid.Redraw
    
     
    With vsGrid
        .Redraw = flexRDNone
        If lngRow >= 0 And lngRow < .Rows - 1 Then
            i = lngRow
        Else
            If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) <> "" Then .Rows = .Rows + 1
            i = .Rows - 1
        End If
        
        For Each objBalanceItem In objBalanceItems
            
            Set objCard = objBalanceItem.objCard
            If objCard Is Nothing Then
                'zlGetCard(ByVal lngCardTypeId As Long, ByVal bln���ѿ� As Boolean,    ByRef objCard As Card)
                If objBalanceItem.�����ID <> 0 Then
                    Call gobjSquare.objOneCardComLib.zlGetCard(objBalanceItem.�����ID, objBalanceItem.���ѿ�, objCard)
                Else
                    Set objCard = zlGetCardFromBalanceName(objBalanceItem.���㷽ʽ)
                End If
                Set objBalanceItem.objCard = objCard
            End If
            
            If objCard Is Nothing Then
                Set objCard = New Card
                With objCard
                    .���㷽ʽ = objBalanceItem.���㷽ʽ
                    .�������� = objBalanceItem.��������
                    .�Ƿ����� = True
                    .�Ƿ�ȫ�� = False
                    .���� = ""
                End With
                Set objBalanceItem.objCard = objCard
            End If
            
            .TextMatrix(i, .ColIndex("����")) = objBalanceItem.��������
            .TextMatrix(i, .ColIndex("�����ID")) = objBalanceItem.�����ID
            .TextMatrix(i, .ColIndex("���ѿ�ID")) = objBalanceItem.���ѿ�ID
            .TextMatrix(i, .ColIndex("��������")) = objBalanceItem.��������
            .TextMatrix(i, .ColIndex("�༭״̬")) = IIf(objBalanceItem.�Ƿ�����༭, "1", "0") & "|" & IIf(objBalanceItem.�Ƿ�����ɾ��, "1", "0")      '�Ƿ�����༭|�Ƿ�����ɾ��
            .TextMatrix(i, .ColIndex("�Ƿ�����")) = IIf(objCard.�Ƿ�����, 1, 0)
            .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = IIf(objCard.�Ƿ�ȫ��, 1, 0)
            .TextMatrix(i, .ColIndex("У�Ա�־")) = objBalanceItem.У�Ա�־
            .TextMatrix(i, .ColIndex("�Ƿ�����")) = IIf(objBalanceItem.�Ƿ�����, 1, 0)
            .TextMatrix(i, .ColIndex("���������")) = objCard.����
            .TextMatrix(i, .ColIndex("���㷽ʽ")) = objBalanceItem.���㷽ʽ
            .TextMatrix(i, .ColIndex("������")) = IIf(objBalanceItem.�������� = 9, Format(objBalanceItem.������, "###0.00#####"), Format(objBalanceItem.������, "0.00"))
            .TextMatrix(i, .ColIndex("�������")) = objBalanceItem.�������
            .TextMatrix(i, .ColIndex("��ע")) = objBalanceItem.����ժҪ
            .TextMatrix(i, .ColIndex("������ˮ��")) = objBalanceItem.������ˮ��
            .TextMatrix(i, .ColIndex("����˵��")) = objBalanceItem.����˵��
            .TextMatrix(i, .ColIndex("ԭԤ��id")) = objBalanceItem.Ԥ��ID
            .TextMatrix(i, .ColIndex("����")) = IIf(objBalanceItem.�Ƿ�����, String(Len(objBalanceItem.����), "*"), objBalanceItem.����)
            .Cell(flexcpData, i, .ColIndex("����")) = NVL(objBalanceItem.����)
            .RowData(i) = objBalanceItem
            If bln�鿴 Then
                If objBalanceItem.У�Ա�־ = 1 Then    'δִ�гɹ���
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                ElseIf objBalanceItem.У�Ա�־ = 2 Then  'ִ�гɹ��ҵ�ǰ���ڲ鿴��
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                Else
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vsGrid.ForeColor
                End If
            End If
            
            If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) <> "" Then .Rows = .Rows + 1
            i = .Rows - 1
        Next
        .Redraw = bytPreDraw
    End With
    zlLoadBalanceItemsToVsGrid = True
    Exit Function
errHandle:
    vsGrid.Redraw = bytPreDraw
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Sub zlClearSquareBalance(ByVal vsBalance As VSFlexGrid, ByVal lngCardTypeID As Long, _
     ByRef objBalanceInfor As clsBalanceInfo, Optional ByVal lng���ѿ�ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ѿ�����
    '����:���˺�
    '����:2015-01-23 14:54:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, j As Long
    With vsBalance
        j = 1
        Do While j <= .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("����"))) = 5 _
                And Val(.TextMatrix(j, .ColIndex("�����ID"))) = lngCardTypeID _
                And (lng���ѿ�ID = 0 Or (lng���ѿ�ID <> 0 And Val(.TextMatrix(j, .ColIndex("���ѿ�ID"))) = lng���ѿ�ID)) Then
                dblMoney = Val(.TextMatrix(j, .ColIndex("������")))
                
                objBalanceInfor.�Ѹ��ϼ� = RoundEx(objBalanceInfor.�Ѹ��ϼ� - dblMoney, 6)
                objBalanceInfor.δ���ϼ� = RoundEx(objBalanceInfor.δ���ϼ� + dblMoney, 6)
                If .Rows >= 2 Then
                    .RemoveItem j
                Else
                    .Rows = 2
                   .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                   .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                   .RowData(1) = ""
                   j = 2
                End If
            Else
                j = j + 1
            End If
        Loop
    End With
End Sub
Public Function zlGetBalanceNULLRow(ByVal vsBalance As VSFlexGrid, Optional lngRow As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㷽ʽ��ΪNULL�Ľ��㷽ʽ����
    '���:lngRow-��ǰ�к������
    '����:
    '����:-1��ʾ������;>1 ��ʾ��ȡ�ɹ�����
    '����:���˺�
    '����:2018-04-10 14:18:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If lngRow = 0 Then lngRow = 1
    With vsBalance
        For i = lngRow To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("���㷽ʽ"))) = "" Then
                zlGetBalanceNULLRow = i: Exit Function
            End If
        Next
    End With
    zlGetBalanceNULLRow = -1
End Function

Public Sub zlRecalItemObjectRowNo(ByVal vsBalance As VSFlexGrid)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ��Item������к�(�Ա����к�����ȷ��)
    '����:���˺�
    '����:2018-07-13 14:06:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItemTemp As clsBalanceItem
    On Error GoTo errHandle
    With vsBalance
         For i = 1 To .Rows - 1
             If zlGetBalanceItemFromBalanceGrid(vsBalance, i, objItemTemp) Then
                objItemTemp.�к� = i
                .RowData(i) = objItemTemp
             End If
         Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlCheckVsBalanceIsExsitsFromCardObject(ByVal vsBalance As VSFlexGrid, ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���Ŀ����󣬼��ָ���Ľ����б����Ƿ����
    '���:objCard-������
    '����:
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2018-07-24 14:54:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln���ѿ� As Boolean, i As Long
    On Error GoTo errHandle
    If objCard Is Nothing Then Exit Function
    
    With vsBalance
        For i = 1 To .Rows - 1
            If objCard.�ӿ���� > 0 Then
                bln���ѿ� = Val(.TextMatrix(i, .ColIndex("����"))) = 5
                If objCard.�ӿ���� = Val(.TextMatrix(i, .ColIndex("�����ID"))) And objCard.���ѿ� = bln���ѿ� Then
                     zlCheckVsBalanceIsExsitsFromCardObject = True: Exit Function
                End If
                
            Else
                If objCard.���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ"))) Then
                    zlCheckVsBalanceIsExsitsFromCardObject = True: Exit Function
                End If
            End If
        Next i
    End With
    zlCheckVsBalanceIsExsitsFromCardObject = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlInitDepositGrid(ByVal vsDeposit As VSFlexGrid, ByVal lngModul As Long, ByVal strFromName As String, ByVal strRegName As String, _
    Optional bytEditType As gBalanceBill, Optional blnAllowSort As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-12-29 15:08:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsDeposit
        .Clear
        .Cols = 26: .Rows = 2
        i = 0
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "���ݺ�": i = i + 1
        .TextMatrix(0, i) = "���": i = i + 1
        .TextMatrix(0, i) = "Ʊ�ݺ�": i = i + 1
        .TextMatrix(0, i) = "�տ�����": i = i + 1
        .TextMatrix(0, i) = "���㷽ʽ": i = i + 1
        .TextMatrix(0, i) = "���": i = i + 1
        .TextMatrix(0, i) = "��Ԥ��": i = i + 1
        .TextMatrix(0, i) = "���": i = i + 1
        .TextMatrix(0, i) = "Ԥ��ID": i = i + 1
        .TextMatrix(0, i) = "�༭״̬": i = i + 1
        .TextMatrix(0, i) = "�����ID": i = i + 1
        .TextMatrix(0, i) = "�Ƿ����ѿ�": i = i + 1
        .TextMatrix(0, i) = "���������": i = i + 1
        .TextMatrix(0, i) = "����": i = i + 1
        .TextMatrix(0, i) = "������ˮ��": i = i + 1
        .TextMatrix(0, i) = "����˵��": i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ȫ��": i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ȱʡ����": i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ת�ʼ�����": i = i + 1
        .TextMatrix(0, i) = "��������ID": i = i + 1
        .TextMatrix(0, i) = "��������": i = i + 1
        .TextMatrix(0, i) = "ԭʼ���": i = i + 1
        .TextMatrix(0, i) = "�������": i = i + 1
        .TextMatrix(0, i) = "ժҪ": i = i + 1
          
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedCols = 1
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            
            ''ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)|������(0-��������,1-��ֹ����,2-��������,�����س���������)
            Select Case .ColKey(i)
            Case "���ݺ�"
                .ColData(i) = "1|0"
                .FixedAlignment(i) = flexAlignRightCenter
            Case "���"
                 If bytEditType = g_Ed_������� Or bytEditType = g_Ed_סԺ���� _
                    Or bytEditType = g_Ed_���½��� Then
                    .ColData(i) = "0|0"
                    .ColHidden(i) = False
                 Else
                      .ColHidden(i) = True: .ColData(i) = "-1|1"
                 End If
            Case "��Ԥ��"
                    .ColData(i) = "1|0"
                    .ColHidden(i) = False
            Case "���"
                 If bytEditType = g_Ed_������� Or bytEditType = g_Ed_סԺ���� Or bytEditType = g_Ed_���½��� Then
                     .ColHidden(i) = True: .ColData(i) = "0|1"
                 Else
                      .ColHidden(i) = True: .ColData(i) = "-1|0"
                 End If
            Case "���������", "����", "������ˮ��", "����˵��", "�������", "ժҪ"
                 .ColHidden(i) = True: .ColData(i) = "0|0"
            Case Else
                If Not .ColKey(i) Like "*ID" Then
                    .ColData(i) = "0|0"
                End If
            End Select
            
            If InStr(",�Ƿ����ѿ�,�Ƿ�����,�Ƿ�ȫ��,�Ƿ�ȱʡ����,�Ƿ�ת�ʼ�����,�༭״̬,��������,ԭʼ���,", "," & .ColKey(i) & ",") > 0 Or .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"
            ElseIf .ColKey(i) Like "*��" Or .ColKey(i) Like "*��Ԥ��" Then
                .ColAlignment(i) = flexAlignRightCenter
            End If
        Next
        
        .ExtendLastCol = False
        .ExplorerBar = IIf(blnAllowSort, flexExSort, flexExNone)
        .ColHidden(.ColIndex("���")) = True
        .ColWidth(.ColIndex("���")) = 1100
        
        .ColHidden(.ColIndex("Ʊ�ݺ�")) = True
        .ColWidth(.ColIndex("Ʊ�ݺ�")) = 1100
        .ColWidth(.ColIndex("�տ�����")) = 1200
        .ColWidth(.ColIndex("���ݺ�")) = 1100
        .ColWidth(.ColIndex("���㷽ʽ")) = 1400
        .ColWidth(.ColIndex("���")) = 1100
        .ColWidth(.ColIndex("��Ԥ��")) = 1100
        .ColWidth(.ColIndex("���������")) = 1800
        .ColWidth(.ColIndex("����")) = 1100
        .ColWidth(.ColIndex("������ˮ��")) = 1100
        .ColWidth(.ColIndex("����˵��")) = 1600
        .ColWidth(.ColIndex("�������")) = 1100
        .ColWidth(.ColIndex("ժҪ")) = 1600
        .ColHidden(.ColIndex("���")) = True
        .ColData(.ColIndex("���")) = "-1|1"
        zl_vsGrid_Para_Restore lngModul, vsDeposit, strFromName, strRegName
        If bytEditType = g_Ed_���ݲ鿴 Or bytEditType = g_Ed_�������� Or bytEditType = g_Ed_�������� Or bytEditType = g_Ed_ȡ������ Then
            .ColHidden(.ColIndex("���")) = True: .ColData(.ColIndex("���")) = "-1|1"
        End If
    End With
    zlInitDepositGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Sub zlInitBalanceGrid(ByVal vsBalance As VSFlexGrid, ByVal lngModul As Long, ByVal strFromName As String, ByVal strRegKey As String, Optional bln�鿴 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�������б�
    '����:���˺�
    '����:2015-01-23 14:14:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBalance
    
        For i = 1 To .Rows - 1
            .RowData(i) = ""
        Next
        .Clear: .Rows = 2: i = 0: .Cols = 21
        .TextMatrix(0, i) = "�����ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "���ѿ�ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "��������": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�༭״̬": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "����״̬": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ȫ��": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "У�Ա�־": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": .ColWidth(i) = 0: i = i + 1
        
        .TextMatrix(0, i) = "���㷽ʽ": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "������": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "�������": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "���������": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "������ˮ��": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "����˵��": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "��ע": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "��������ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "ԭԤ��ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ת��": .ColWidth(i) = 0: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
            Select Case .ColKey(i)
            Case "�Ƿ�ת��", "��������ID", "��������", "����", "�Ƿ񱣴�", "�Ƿ�����", "У�Ա�־", "�༭״̬", "�Ƿ�����", "�Ƿ�ȫ��", "����״̬", "�Ƿ���֤", "ԭԤ��ID"
                .ColHidden(i) = True
                .ColData(i) = """-1||1"
            Case "������"
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = """1||0"
            Case .ColIndex("���㷽ʽ")
                .ColData(i) = """1||0"
            Case "���������"
                .ColData(i) = "1||2"
            Case .ColIndex("�������")
                .ColData(i) = "1||0"
            Case Else
                .ColData(i) = "1||" & IIf(bln�鿴, "0", "2")
                
            End Select
            If bln�鿴 Then .ColData(i) = ""
        Next
        If Not bln�鿴 Then .Editable = flexEDKbdMouse
    End With
    zl_vsGrid_Para_Restore lngModul, vsBalance, strFromName, strRegKey
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



Public Sub zlAutoRecalFeeBalanceMoney(ByVal vsDetailList As VSFlexGrid, ByVal dbl���ν��� As Double, ByVal dbl���ο�� As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Զ�����ͷ�̯���ʽ��
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-24 15:59:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Long
    Dim blnAll As Boolean
    
    On Error GoTo errHandle
    
    dblMoney = dbl���ν���
    blnAll = dbl���ν��� = dbl���ο��
    With vsDetailList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����")) <> "" Then
                If dblMoney >= Val(.Cell(flexcpData, i, .ColIndex("δ����"))) And dblMoney <> 0 Or blnAll Then
                    .Cell(flexcpData, i, .ColIndex("���ʽ��")) = Val(.Cell(flexcpData, i, .ColIndex("δ����")))
                    dblMoney = RoundEx(dblMoney - Val(.Cell(flexcpData, i, .ColIndex("���ʽ��"))), 6)
                Else
                    If dblMoney = 0 Then
                        .Cell(flexcpData, i, .ColIndex("���ʽ��")) = ""
                    Else
                        .Cell(flexcpData, i, .ColIndex("���ʽ��")) = dblMoney
                    End If
                    dblMoney = 0
                End If
                .TextMatrix(i, .ColIndex("���ʽ��")) = Format(Val(.Cell(flexcpData, i, .ColIndex("���ʽ��"))), gstrDec)
            End If
        Next i
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Function zlGetRemainderMoneyToPati(ByVal byt���� As Byte, ByVal lng����ID As Long, ByRef objPati As clsPatiInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����������������˶���
    '���:objPati-��ǰ�Ĳ���
    '     byt����-1-����;2-סԺ;0-���з������
    '����:objPati-���ظ��µĲ��������Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-24 20:16:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If byt���� = 0 Then
        strSQL = "Select sum(Ԥ�����) As Ԥ�����,sum(�������) As ������� From ������� Where ����ID= [1] And ����=1"
    Else
        strSQL = "Select sum(Ԥ�����) as Ԥ�����,sum(�������) as ������� From ������� Where ����ID= [1] And ����=1 And ����= [2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���������Ϣ", lng����ID, byt����)
    
    objPati.Ԥ����� = Format(Val(NVL(rsTemp!Ԥ�����)), "0.00")
    objPati.������� = Format(Val(NVL(rsTemp!�������)), "0.00")
    objPati.Ԥ��ʣ��ϼ� = Format(Val(NVL(rsTemp!Ԥ�����)) - Val(NVL(rsTemp!�������)), "0.00")
    zlGetRemainderMoneyToPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckNoSettlementMoney(ByVal str���� As String, _
    ByVal lng����ID As Long, ByVal strTimes As String, _
    Optional ByVal byt�������� As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������۲����Ƿ����δ����ý��
    '���:
    '   lng����ID ָ������
    '   strTimes ָ��סԺ����,�����Ӣ�Ķ��ŷָ���Ϊ�ձ�ʾ����סԺ����
    '   byt�������� 1-�������;2-סԺ����
    '����:
    '����:���ͨ������True,���򷵻�False
    '˵��:����������۲��ˣ�����סԺ����ʱ��������������������ʾ�������������ʱ���������סԺ����������Ƚ�סԺ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strWhere As String
    Dim strTemp As String
    
    On Error GoTo ErrHandler
    If strTimes <> "" Then
        strWhere = " And a.��ҳid In(Select /*+Cardinality(j,10)*/ Column_Value From Table(f_Num2list([3])) J)"
    End If
    strSQL = "Select a.��ҳID,Sum(a.���) As δ����" & _
            " From ����δ����� A" & _
            " Where a.����id=[1] And a.��Դ;�� = [2]" & strWhere & _
            " Group By a.��ҳID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����δ�����", lng����ID, IIf(byt�������� = 1, 2, 1), strTimes)
    If rsTemp.EOF Then zlCheckNoSettlementMoney = True: Exit Function
    
    Do While Not rsTemp.EOF
        strTemp = strTemp & "," & lng����ID & ":" & Val(NVL(rsTemp!��ҳID))
        rsTemp.MoveNext
    Loop
    strTemp = Mid(strTemp, 2)
    
    '����Ƿ�Ϊ��������סԺ
    If zlGetPatiPageInfo(0, strTemp, rsTemp) = False Then Exit Function
    rsTemp.Filter = "��������=1"
    If rsTemp.EOF Then zlCheckNoSettlementMoney = True: Exit Function
    
    Do While Not rsTemp.EOF
        strTemp = strTemp & "��" & Val(NVL(rsTemp!��ҳID))
        rsTemp.MoveNext
    Loop
    strTemp = Mid(strTemp, 2)
    
    If byt�������� = 2 Then
        MsgBox "���ˡ�" & str���� & "���ڵ�" & strTemp & "��סԺ������δ�����������ã�ע��������������ʣ�", vbInformation, gstrSysName
    Else
        MsgBox "���ˡ�" & str���� & "���ڵ�" & strTemp & "���������ۻ�����δ�����סԺ���ã������ȶ������סԺ���ˣ�", vbInformation, gstrSysName
        Exit Function
    End If
    zlCheckNoSettlementMoney = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlLoadDetaiFeeToGridFromRecord(ByVal rsFeeList As ADODB.Recordset, ByVal bln���� As Boolean, ByVal intInsure As Integer, ByRef vsDetailList As VSFlexGrid, _
    ByVal lngModule As Long, ByVal strFromName As String, strRegKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�Ŀ������
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-05 18:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
  
    On Error GoTo errHandle
    If rsFeeList Is Nothing Then Exit Function
    If rsFeeList.State <> 1 Then Exit Function
    
    If rsFeeList.RecordCount <> 0 Then rsFeeList.MoveFirst
    With vsDetailList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        Do While Not rsFeeList.EOF
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Format(NVL(rsFeeList!ʱ��), "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(rsFeeList!���ݺ�)
            .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = NVL(rsFeeList!��Ŀ)
            .TextMatrix(.Rows - 1, .ColIndex("δ����")) = Format(NVL(rsFeeList!δ����), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("δ����")) = Val(NVL(rsFeeList!δ����))
            .TextMatrix(.Rows - 1, .ColIndex("���ʽ��")) = Format(NVL(rsFeeList!���ʽ��), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("���ʽ��")) = Val(NVL(rsFeeList!���ʽ��))
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsFeeList!ID, 0)
            .TextMatrix(.Rows - 1, .ColIndex("��¼����")) = Val(NVL(rsFeeList!��¼����))
            .TextMatrix(.Rows - 1, .ColIndex("��¼״̬")) = IIf(Val(NVL(rsFeeList!��¼״̬)) = 3, 1, Val(NVL(rsFeeList!��¼״̬)))
            .TextMatrix(.Rows - 1, .ColIndex("ִ��״̬")) = Val(NVL(rsFeeList!ִ��״̬))
            .TextMatrix(.Rows - 1, .ColIndex("���")) = Val(NVL(rsFeeList!���))
            If bln���� Then .Cell(flexcpData, .Rows - 1, .ColIndex("���")) = Val(NVL(rsFeeList!�����־))
            .Rows = .Rows + 1
            rsFeeList.MoveNext
        Loop
        .Cell(flexcpBackColor, 1, .ColIndex("���ʽ��"), .Rows - 1, .ColIndex("���ʽ��")) = IIf(intInsure <> 0, .Cell(flexcpBackColor, 1, .ColIndex("����")), &HFFFFC0)
        If .TextMatrix(1, .ColIndex("����")) <> "" Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore lngModule, vsDetailList, strFromName, strRegKey

    zlLoadDetaiFeeToGridFromRecord = True
    Exit Function
errHandle:
     vsDetailList.Redraw = flexRDNone
     Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlLoadDetaiFeeToGridFromBalanceID(ByVal lng����ID As Long, ByVal vsDetailList As VSFlexGrid, _
    ByVal bln������� As Boolean, ByVal bln���� As Boolean, ByVal lngModule As Long, ByVal strFromName As String, strRegKey As String, Optional blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID�����ط�Ŀ������
    '���:lng����ID-����ID
    '     blnNoMoved-�Ƿ���ʷ����ת��
    '     bln����-�Ƿ�鿴���ϼ�¼
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 18:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, lngRow As Long, intSign As Integer
    
    On Error GoTo errHandle
    
    intSign = IIf(bln����, -1, 1)
    
    strSQL = _
    "   Select Mod(A.��¼����,10) as ��¼����, A.NO,A.���," & _
    "          Max(Decode(a.�Ƿ���,1,'***',b.����)) as ��Ŀ," & _
    "          Max(A.����ʱ��) As ����ʱ��, " & _
    "          Sum(Round(A.��׼����*A.����*Nvl(A.����,1)," & gbytDec & ")) as ��׼���,Sum(A.���ʽ��) as ���ʽ��, " & _
    "          Decode(A.��¼״̬,2,2,1) As ��¼״̬,Max(a.�����־) As �����־ " & _
    "   From סԺ���ü�¼ A,�շ���ĿĿ¼ B" & _
    "   Where A.����ID= [1] And A.�շ�ϸĿID=B.ID " & _
    "   Group by Mod(A.��¼����,10),A.NO,A.���,Decode(A.��¼״̬,2,2,1) "
   
    strSQL = strSQL & " UNION ALL " & vbCrLf & _
        Replace(strSQL, "סԺ���ü�¼", "������ü�¼")

    If blnNOMoved Then
        strSQL = Replace(Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼"), "������ü�¼", "H������ü�¼")
    End If
    
    strSQL = "" & _
    "   Select Max(����ʱ��) As ����ʱ��,NO,���,��Ŀ, sum(��׼���) as ��׼���," & _
    "          sum(���ʽ��) as ���ʽ��,��¼״̬,Max(�����־) As �����־ " & _
    "   From (" & strSQL & ")" & _
    "   Group by NO,���,��Ŀ,��¼״̬" & _
    "   Order by NO,���"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݽ���IDͳ�Ʒ�����ϸ", lng����ID)
    
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsDetailList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        
        .TextMatrix(0, .ColIndex("δ����")) = "��׼���"
        .TextMatrix(0, .ColIndex("���ʽ��")) = IIf(intSign = -1, "���Ͻ��", "���ʽ��")
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Format(NVL(rsTemp!����ʱ��), "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(rsTemp!NO)
            .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = NVL(rsTemp!��Ŀ)
            .TextMatrix(.Rows - 1, .ColIndex("δ����")) = Format(NVL(rsTemp!��׼���), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("δ����")) = Val(NVL(rsTemp!��׼���))
            .TextMatrix(.Rows - 1, .ColIndex("���ʽ��")) = Format(intSign * Val(NVL(rsTemp!���ʽ��)), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("���ʽ��")) = intSign * Val(NVL(rsTemp!���ʽ��))
            .TextMatrix(.Rows - 1, .ColIndex("���")) = Val(NVL(rsTemp!���))
            If bln������� Then
                .Cell(flexcpData, .Rows - 1, .ColIndex("���")) = Val(NVL(rsTemp!�����־))
            End If
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        .Cell(flexcpBackColor, 1, .ColIndex("���ʽ��"), .Rows - 1, .ColIndex("���ʽ��")) = .Cell(flexcpBackColor, 1, .ColIndex("����"), 0.1, .ColIndex("����"))
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore lngModule, vsDetailList, strFromName, strRegKey
    zlLoadDetaiFeeToGridFromBalanceID = True
    Exit Function
errHandle:
    vsDetailList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

 
Public Function zlLoadFeiMuFeeListToGridFromRecord(ByVal rsFeeList As ADODB.Recordset, ByVal bln���� As Boolean, ByVal intInsure As Integer, ByRef vsFeeList As VSFlexGrid, _
    ByVal lngModule As Long, ByVal strFromName As String, ByVal strRegKey As String, ByRef dblMoney_out As Double) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�Ŀ������
    '���:rsFeeList-���ü�
    '     bln����-�Ƿ��������
    '     intInsure-����
    '����:dblMoney_out-����δ����ϼ�
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-05 18:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dblMoney(0 To 2) As Double
   
    On Error GoTo errHandle
    
    dblMoney_out = 0
    If rsFeeList Is Nothing Then Exit Function
    If rsFeeList.State <> 1 Then Exit Function

    If rsFeeList.RecordCount <> 0 Then rsFeeList.MoveFirst
     With vsFeeList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        Do While Not rsFeeList.EOF
           lngRow = .FindRow(NVL(rsFeeList!��Ŀ, "δ֪"), "1", .ColIndex("��Ŀ"), , True)
           If lngRow < 0 Then
                If .TextMatrix(1, .ColIndex("��Ŀ")) = "" Then
                    lngRow = 1
                Else
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                End If
           End If
           
           If .TextMatrix(1, .ColIndex("��Ŀ")) = "" Then lngRow = 1
          .TextMatrix(lngRow, .ColIndex("��Ŀ")) = NVL(rsFeeList!��Ŀ, "δ֪")
          
          .Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��")) = Val(.Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��"))) + Val(NVL(rsFeeList!Ӧ�ս��))
          .TextMatrix(lngRow, .ColIndex("Ӧ�ս��")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��")) = Val(.Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��"))) + Val(NVL(rsFeeList!ʵ�ս��))
          .TextMatrix(lngRow, .ColIndex("ʵ�ս��")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("δ����")) = Val(.Cell(flexcpData, lngRow, .ColIndex("δ����"))) + Val(NVL(rsFeeList!δ����))
          .TextMatrix(lngRow, .ColIndex("δ����")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("δ����"))), gstrDec)
            
          dblMoney(0) = RoundEx(dblMoney(0) + Val(NVL(rsFeeList!Ӧ�ս��)), 5)
          dblMoney(1) = RoundEx(dblMoney(1) + Val(NVL(rsFeeList!ʵ�ս��)), 5)
          dblMoney(2) = RoundEx(dblMoney(2) + Val(NVL(rsFeeList!δ����)), 5)
          
            rsFeeList.MoveNext
        Loop
        
        .ColSort(.ColIndex("��Ŀ")) = flexSortUseColSort
        If .TextMatrix(1, .ColIndex("��Ŀ")) <> "" Then
          .Rows = .Rows + 1: lngRow = .Rows - 1
          .TextMatrix(lngRow, .ColIndex("��Ŀ")) = "�ϼ�"
          .Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��")) = dblMoney(0)
          .TextMatrix(lngRow, .ColIndex("Ӧ�ս��")) = Format(dblMoney(0), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��")) = dblMoney(1)
          .TextMatrix(lngRow, .ColIndex("ʵ�ս��")) = Format(dblMoney(1), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("δ����")) = dblMoney(2)
          .TextMatrix(lngRow, .ColIndex("δ����")) = Format(dblMoney(2), gstrDec)
         
           .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
    End With
    
    zl_vsGrid_Para_Restore lngModule, vsFeeList, strFromName, strRegKey
    dblMoney_out = dblMoney(2)
    zlLoadFeiMuFeeListToGridFromRecord = True
    Exit Function
errHandle:
     vsFeeList.Redraw = flexRDNone
     Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlLoadFeiMuFeeListToGridFromBalanceID(ByVal lng����ID As Long, ByVal vsFeeList As VSFlexGrid, _
    ByVal lngModule As Long, ByVal bln���� As Boolean, ByRef dblBalanceMoney_Out As Double, ByVal strFromName As String, strRegKey As String, Optional blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID,���վݷ�Ŀͳ�Ʒ��ú󣬽����ݼ��ص�����
    '���:lng����ID-����ID
    '    vsFeeList-�����б�����
    '����:dblBalanceMoney_Out-��ǰ���ʺϼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-25 10:00:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long, strSQL As String
    Dim dblMoney(0 To 1) As Double, intSign As Integer
    
    On Error GoTo errHandle
    
    intSign = IIf(bln����, -1, 1)
    strSQL = "" & _
    "   Select Mod(A.��¼����,10) as ��¼����, A.NO,���,A.�վݷ�Ŀ, " & _
    "          sum(Round(A.��׼����*A.����*Nvl(A.����,1)," & gbytDec & ")) as ��׼���,sum(A.���ʽ��) as ���ʽ�� " & _
    "   From סԺ���ü�¼ A " & _
    "   Where A.����ID= [1]  " & _
    "   Group by Mod(A.��¼����,10),A.NO,A.���,A.�վݷ�Ŀ "
    
   
    strSQL = strSQL & " UNION ALL " & vbCrLf & _
        Replace(strSQL, "סԺ���ü�¼", "������ü�¼")

    If blnNOMoved Then strSQL = Replace(Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼"), "������ü�¼", "H������ü�¼")
    
    strSQL = "" & _
    "   Select �վݷ�Ŀ, sum(��׼���) as ��׼���,sum(���ʽ��) as ���ʽ�� " & _
    "   From (" & strSQL & ")" & _
    "   Group by �վݷ�Ŀ" & _
    "   Order by �վݷ�Ŀ"
    
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݽ���IDͳ�ƽ��ʷ�Ŀ��ϸ", lng����ID)
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    
     dblMoney(0) = 0: dblMoney(1) = 0
    With vsFeeList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        
         lngRow = 1
        
        .TextMatrix(0, .ColIndex("Ӧ�ս��")) = "��׼���"
        .ColHidden(.ColIndex("ʵ�ս��")) = True
        
        Do While Not rsTemp.EOF
          .TextMatrix(lngRow, .ColIndex("��Ŀ")) = NVL(rsTemp!�վݷ�Ŀ, "δ֪")
          .Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��")) = Val(.Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��"))) + Val(NVL(rsTemp!��׼���))
          .TextMatrix(lngRow, .ColIndex("Ӧ�ս��")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��"))), gstrDec)
         
          .Cell(flexcpData, lngRow, .ColIndex("δ����")) = Val(.Cell(flexcpData, lngRow, .ColIndex("���ʽ��"))) + Val(NVL(rsTemp!���ʽ��))
          .TextMatrix(lngRow, .ColIndex("δ����")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("���ʽ��"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("���ʽ��")) = Val(.Cell(flexcpData, lngRow, .ColIndex("���ʽ��"))) + RoundEx(intSign * Val(NVL(rsTemp!���ʽ��)), 6)
          .TextMatrix(lngRow, .ColIndex("���ʽ��")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("���ʽ��"))), gstrDec)
          
          dblMoney(0) = dblMoney(0) + Val(NVL(rsTemp!��׼���))
          dblMoney(1) = dblMoney(1) + RoundEx(intSign * Val(NVL(rsTemp!���ʽ��)), 6)
          .Rows = .Rows + 1: lngRow = .Rows - 1
          rsTemp.MoveNext
        Loop
        dblMoney(0) = RoundEx(dblMoney(0), 5)
        dblMoney(1) = RoundEx(dblMoney(1), 5)
        
        If .TextMatrix(1, .ColIndex("��Ŀ")) <> "" Then
           lngRow = .Rows - 1
          .TextMatrix(lngRow, .ColIndex("��Ŀ")) = "�ϼ�"
          .Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��")) = dblMoney(0)
          .TextMatrix(lngRow, .ColIndex("Ӧ�ս��")) = Format(dblMoney(0), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("δ����")) = dblMoney(1)
          .TextMatrix(lngRow, .ColIndex("δ����")) = Format(dblMoney(1), gstrDec)
         
          .Cell(flexcpData, lngRow, .ColIndex("���ʽ��")) = dblMoney(1)
          .TextMatrix(lngRow, .ColIndex("���ʽ��")) = Format(dblMoney(1), gstrDec)
         
           .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True
        End If
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
    End With
    
    dblBalanceMoney_Out = dblMoney(1)
    zl_vsGrid_Para_Restore lngModule, vsFeeList, strFromName, strRegKey
    
    zlLoadFeiMuFeeListToGridFromBalanceID = True
    Exit Function
errHandle:
    vsFeeList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function zlLoadDepositListFromBalanceID(ByVal lng����ID As Long, vsDeposit As VSFlexGrid, ByVal blnNOMoved As Boolean, _
    ByRef dblTotal_Out As Double, ByRef rsDeposit_Out As ADODB.Recordset, ByRef intCountBill_Out As Integer, _
    ByVal lngModul As Long, Optional strFormName As String, Optional strRegKey As String, Optional bln���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID��ȡ��Ԥ����Ϣ��Ϣ�����ص�Ԥ���б���
    '���:lng����ID-ָ���Ľ���ID
    '     blnNoMoved-��ǰ�Ƿ��ƶ����󱸱���
    '     bln����-�Ƿ�鿴���ϵ���
    '����:rsDeposit_Out-����Ԥ����¼��
    '     dblTotal_Out-��Ԥ���ܼ�
    '     intCountBill_Out-�漰��Ʊ������
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 15:09:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, dblTotal As Double
    Dim intSign As Integer
    
    On Error GoTo errHandle
    dblTotal_Out = 0
    Set rsTemp = GetBalanceDeposit(lng����ID, blnNOMoved)
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    intSign = IIf(bln����, -1, 1)
    With vsDeposit
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        'ID,���ݺ�,Ʊ�ݺ�,����,���㷽ʽ, ���
        i = 1
        Do While Not rsTemp.EOF
            .RowData(i) = ""
            .TextMatrix(i, .ColIndex("ID")) = rsTemp!ID
            .TextMatrix(i, .ColIndex("���ݺ�")) = rsTemp!���ݺ�
            .TextMatrix(i, .ColIndex("���")) = NVL(rsTemp!Ԥ�����)
            .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) = "" & rsTemp!Ʊ�ݺ�
            .TextMatrix(i, .ColIndex("�տ�����")) = Format(rsTemp!����, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("���㷽ʽ")) = NVL(rsTemp!���㷽ʽ)
            .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(intSign * rsTemp!���, "0.00")
            .TextMatrix(i, .ColIndex("�����ID")) = Val(NVL(rsTemp!�����ID))
            .TextMatrix(i, .ColIndex("�Ƿ����ѿ�")) = Val(NVL(rsTemp!�Ƿ����ѿ�))
            .TextMatrix(i, .ColIndex("���������")) = NVL(rsTemp!���������)
            .TextMatrix(i, .ColIndex("������ˮ��")) = NVL(rsTemp!������ˮ��)
            .TextMatrix(i, .ColIndex("�������")) = NVL(rsTemp!�������)
            .TextMatrix(i, .ColIndex("ժҪ")) = NVL(rsTemp!ժҪ)
            .TextMatrix(i, .ColIndex("����")) = NVL(rsTemp!����)
            .TextMatrix(i, .ColIndex("Ԥ��ID")) = Val(NVL(rsTemp!ID))
            .TextMatrix(i, .ColIndex("����˵��")) = NVL(rsTemp!����˵��)
            .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(NVL(rsTemp!�Ƿ�����))
            .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = Val(NVL(rsTemp!�Ƿ�ȫ��))
            .TextMatrix(i, .ColIndex("�Ƿ�ȱʡ����")) = Val(NVL(rsTemp!�Ƿ�ȱʡ����))
            .TextMatrix(i, .ColIndex("�Ƿ�ת�ʼ�����")) = Val(NVL(rsTemp!�Ƿ�ת�ʼ�����))
            .TextMatrix(i, .ColIndex("��������")) = Val(NVL(rsTemp!��������))
            .TextMatrix(i, .ColIndex("ԭʼ���")) = Val(NVL(rsTemp!ԭʼ���))
            
            .Rows = .Rows + 1: i = i + 1
            dblTotal = dblTotal + intSign * Val(NVL(rsTemp!���))
            rsTemp.MoveNext
        Loop
        .Row = 1: .Col = .Cols - 1
        If i > 1 Then .Rows = .Rows - 1
        
        .ColWidth(.ColIndex("�տ�����")) = 1305
        .ColWidth(.ColIndex("���ݺ�")) = 1100
        .ColWidth(.ColIndex("���㷽ʽ")) = 1400
        .ColWidth(.ColIndex("���")) = 1100
        .ColWidth(.ColIndex("��Ԥ��")) = 1100
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
    End With
    
    zl_vsGrid_Para_Restore lngModul, vsDeposit, strFormName, strRegKey
    
    dblTotal_Out = dblTotal
    intCountBill_Out = rsTemp.RecordCount
    
    Set rsDeposit_Out = rsTemp
    zlLoadDepositListFromBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetDefaultHospitalizedDate(ByVal lng����ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID��ȡ�ϴ���;����ʱ��
    '���:lng����ID-����ID
    '����:�����ϴ���;���ʵĽ�������,����;����ʱ,���ؿ�
    '����:���˺�
    '����:2015-01-06 15:25:02
    '˵��:ԭ�������30043
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select to_char( Max(��������) + 1,'yyyy-mm-dd') as �������� " & _
    "   From ���˽��ʼ�¼ " & _
    "   Where  ��¼״̬=1  And ����iD=[1] and nvl(��;����,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݲ���ID��ȡ�ϴ���;����ʱ��", lng����ID)
    If rsTemp.EOF Then Exit Function
    zlGetDefaultHospitalizedDate = NVL(rsTemp!��������)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsCheck�����ѽ���(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡���Ƿ��Ѿ�����
    '���:
    '����:
    '����:�ѽ��շ���True,���򷵻�False
    '˵��:30036
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    On Error GoTo ErrHandler
    strValue = zlGetPatiPageExtendInfo(lng����ID, lng��ҳID, "��������")
    zlIsCheck�����ѽ��� = Val(strValue) = 1
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlSetBalanceRowDataFromItemsObject(ByVal vsBlance As VSFlexGrid, ByVal objItems As clsBalanceItems)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������Ľ���״̬
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-04-16 19:35:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    On Error GoTo errHandle
    
    If objItems Is Nothing Then Exit Sub
    
    For Each objItem In objItems
        Call zlSetBalanceRowDataFromItemObject(vsBlance, objItem)
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlSetBalanceRowDataFromItemObject(ByVal vsBlance As VSFlexGrid, ByVal objItem As clsBalanceItem)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ������ý�������
    '���:objItem-ָ����
    '����:���˺�
    '����:2018-07-13 13:51:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItemTemp As clsBalanceItem
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    If objItem Is Nothing Then Exit Sub
    
    lngRow = objItem.�к�
    
    
    If lngRow > vsBlance.Rows - 1 Or lngRow < 1 Then Exit Sub
    
    If Not zlGetBalanceItemFromBalanceGrid(vsBlance, lngRow, objItemTemp) Then Exit Sub
    
    With vsBlance
        If objItemTemp.�����ID <> objItem.�����ID Then Exit Sub
         
         Set objItemTemp = objItem
        .TextMatrix(lngRow, .ColIndex("�༭״̬")) = IIf(objItemTemp.�Ƿ�����༭, 1, 0) & "|" & IIf(objItemTemp.�Ƿ�����ɾ��, 1, 0)
        .TextMatrix(lngRow, .ColIndex("������")) = Format(objItemTemp.������, "0.00")
        If objItemTemp.�Ƿ���� Then
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbGrayText
        ElseIf objItem.�Ƿ񱣴� Then
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
        Else
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
        End If
        .RowData(lngRow) = objItemTemp
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Function zlGetBalanceCancelSQL(ByRef objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, ByRef cllPro As Collection, _
    Optional blnAllCancel As Boolean, Optional bytУ�Ա�־ As Byte = 0, Optional blnDeleteSQL As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ȡ�����������ϲ��������Sql
    '���:objBalanceInfor-�������
    '     blnAllCancel-�Ƿ�ȫ����
    '     blnDeleteSQL-�Ƿ�ǿ�ƻ�ȡ�˷����SQL
    '����:cllPro-������ر����SQL
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-25 21:00:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng����ID As Long
    Dim i As Long
    
    On Error GoTo errHandle
    
    If cllPro Is Nothing Then Set cllPro = New Collection
    
    If objBalanceInfor.����ID <> 0 Then
        If objBalanceInfor.�Ƿ񱣴���ʵ� And Not blnDeleteSQL And Not blnDeleteSQL Then zlGetBalanceCancelSQL = True: Exit Function
    End If
    
    '��¼���д������Ϲ��̣��������ٴ�����
    For i = 1 To cllPro.Count
        If InStr(UCase(cllPro(i)), UCase("Zl_���˽��ʼ�¼_Cancel")) > 0 Then zlGetBalanceCancelSQL = True: Exit Function
    Next
    
    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    With objBalanceInfor
        .����ID = lng����ID
        .����ʱ�� = zlDatabase.Currentdate
    End With
    
    '���˽����¼������
    strSQL = "Zl_���˽��ʼ�¼_Cancel("
    '  No_In         ���˽��ʼ�¼.No%Type,
    strSQL = strSQL & "'" & objBalanceInfor.���ʵ��ݺ� & "',"
    '  ����id_In     ���˽��ʼ�¼.Id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ����Ա���_In ���˽��ʼ�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ���˽��ʼ�¼.����Ա����%Type
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����ʱ��_In   ���˽��ʼ�¼.�շ�ʱ��%Type := Null
    strSQL = strSQL & "to_date('" & Format(objBalanceInfor.����ʱ��, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')"
    strSQL = strSQL & ")"
    zlAddArray cllPro, strSQL
    
    If Not blnAllCancel Then zlGetBalanceCancelSQL = True: Exit Function
     
    'Zl_���˽�������_Modify
    strSQL = "Zl_���˽�������_Modify("
    '  ��������_In   Number,
    strSQL = strSQL & "" & 0 & ","
    '  ����id_In     ���˽��ʼ�¼.����id%Type,
    strSQL = strSQL & "" & ZVal(objPati.����ID) & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "NULL,"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '�տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
    strSQL = strSQL & "to_date('" & Format(objBalanceInfor.����ʱ��, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '��Ԥ������ids_In Varchar2 := Null,
    ' ����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
    strSQL = strSQL & "NULL,"
    '  �������_In Number:=0
    strSQL = strSQL & "1,"
    '    У�Ա�־_In  Number := 0,
    strSQL = strSQL & "" & bytУ�Ա�־ & ","
    '    ��������id_In    ����Ԥ����¼.Id%Type := Null,
    strSQL = strSQL & "NULL,"
    '    ���ԭ����_In Number:=0
    strSQL = strSQL & "0)"
    zlAddArray cllPro, strSQL
    zlGetBalanceCancelSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetThirdMoneyInforRecordFromSwapID(ByVal str��������IDs As String, ByRef rsSwapRecord_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID,��ȡ��صĽ�������Ϣ��
    '���:str��������IDs-��������ID������ö��ŷ���
    '����:rsSwapRecord_Out-���ع�������ID
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-27 17:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strWhere As String, lng��������ID As Long
    
    On Error GoTo errHandle
    
    If InStr(str��������IDs, ",") > 0 Then
        strWhere = "And A.��������ID In (Select column_value From table(f_num2List([1])) "
        
    Else
       lng��������ID = Val(str��������IDs)
       strWhere = " And  A.��������ID =[2]"
    End If
    
    strSQL = "" & _
    "   Select ��������ID,�����ID,���㷽ʽ,������ˮ��,����˵��, " & vbCrLf & _
    "          nvl(���,0)+decode(mod(��¼����,10),1,0,1)* decode(sign(nvl(��Ԥ��,0)),1,1,0)* nvl(��Ԥ��,0) as ԭʼ���, " & _
    "          decode(sign(nvl(���,0)),-1,1,0)*nvl(���,0)+ decode(sign(nvl(��Ԥ��,0)),-1,1,0)* nvl(��Ԥ��,0) as ���˽��" & _
    "   From ����Ԥ����¼ A " & _
    "   Where 1=1 " & strWhere & _
    "   Union all " & _
    "   Select a.��������ID,a.�����ID,a.���㷽ʽ,a.������ˮ��,a.����˵��, " & vbCrLf & _
    "          0 as ԭʼ���, " & _
    "         -1*nvl(b.���,0) as ���˽��" & _
    "   From ����Ԥ����¼ A,�����˿���Ϣ B" & _
    "   Where  a.ID=b.��¼ID And b.�Ƿ�ת�� =1  " & strWhere

    
    strSQL = "" & _
    " Select ��������ID,�����ID,a.���㷽ʽ,a.������ˮ��,a.����˵��, sum(ԭʼ���) as ԭʼ���, sum(���˽��) as ���˽��, sum(ԭʼ���)-sum(���˽��) as ʣ��δ�˽��" & _
    " From (" & strSQL & ") A " & _
    " Group by a.��������ID,a.�����ID,a.���㷽ʽ,a.������ˮ��,a.����˵��"
    
    Set rsSwapRecord_Out = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������׵�ԭʼ��δ�˽��", str��������IDs, lng��������ID)
    zlGetThirdMoneyInforRecordFromSwapID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlMoveRowBalanceFromSwapID(ByVal vsBalance As VSFlexGrid, _
    lngCardTypeID As Long, ByVal bln���ѿ� As Boolean, ByVal lng��������ID As Long, _
    ByVal lngԤ��ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݹ�������ID,ɾ����Ӧ�Ľ����б��ж�Ӧ����
    '���:vsBalance-����
    '     lngCardTypeID-�����ID
    '     lng��������ID-��������ID
    '     lngԤ��ID-Ԥ�����¼��ID
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-30 11:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As clsBalanceItem
    Dim lngRow As Long
    On Error GoTo errHandle
    With vsBalance
        i = 1
        lngRow = .Row
        Do While i <= .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsBalance, i, objItem) Then
                If objItem.��������ID = lng��������ID And objItem.Ԥ��ID = lngԤ��ID _
                    And objItem.�����ID = lngCardTypeID And objItem.���ѿ� = bln���ѿ� Then
                    .RowData(i) = ""
                    Set objItem = Nothing
                   '��������
                    .RemoveItem i
                Else
                    i = i + 1
                End If
            Else
                i = i + 1
            End If
        Loop
        If .Rows <= 1 Then .Rows = .Rows + 1
        If lngRow > .Rows - 1 Or lngRow <= 1 Then
            .Row = .Rows - 1
        Else
            .Row = lngRow
        End If
    End With
    
    zlMoveRowBalanceFromSwapID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlMoveRowBalanceFromBalanceType(ByVal vsBalance As VSFlexGrid, int���� As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����byt����,ɾ����Ӧ�Ľ����б��ж�Ӧ����
    '���:vsBalance-����
    '     int����-����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    '     lng��������ID-��������ID
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-07-30 11:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As clsBalanceItem
    Dim lngRow As Long, intTYPE As Integer
    On Error GoTo errHandle
    With vsBalance
        i = 1
        lngRow = .Row
        Do While i <= .Rows - 1
            intTYPE = Val(.TextMatrix(i, .ColIndex("����")))
            If int���� = intTYPE Then
                .RowData(i) = ""
                .RemoveItem i   '��������
            Else
                i = i + 1
            End If
        Loop
        If .Rows <= 1 Then .Rows = .Rows + 1
        
        If lngRow > .Rows - 1 Or lngRow <= 1 Then
            .Row = .Rows - 1
        Else
            .Row = lngRow
        End If
    End With
    zlMoveRowBalanceFromBalanceType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlReCalcBalanceInfor(ByVal vsBlance As VSFlexGrid, ByRef objBalanceInfor As clsBalanceInfo, _
    Optional lngNotRow As Long = -1, Optional bln������ As Boolean = True, _
    Optional objCurItem As clsBalanceItem, Optional ByVal bln���� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼��������Ϣ
    '���:objBalanceInfor-��ǰ�Ľ�����Ϣ
    '     objCurItem-��ǰ������Ϣ��(δ�������б���)
    '����:objBalanceInfor-�����Ľ�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-04-11 17:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objItem As clsBalanceItem
    On Error GoTo errHandle
    
    
    objBalanceInfor.δ���ϼ� = 0: objBalanceInfor.�Ѹ��ϼ� = 0
    With vsBlance
        For i = 1 To .Rows - 1
            If i <> lngNotRow Then
                If zlGetBalanceItemFromBalanceGrid(vsBlance, i, objItem) Then
                    If (bln������ And objItem.�������� = 9) Or objItem.�������� <> 9 Then
                        If Not (bln���� And objItem.���ѿ� And objItem.�Ƿ��˿� And objItem.�Ƿ�Ԥ��) Then
                            objBalanceInfor.�Ѹ��ϼ� = RoundEx(objBalanceInfor.�Ѹ��ϼ� + objItem.������, 6)
                        End If
                    End If
                End If
            End If
        Next
    End With
    If Not objCurItem Is Nothing Then
        objBalanceInfor.�Ѹ��ϼ� = RoundEx(objBalanceInfor.�Ѹ��ϼ� + objCurItem.������, 6)
    End If
    objBalanceInfor.δ���ϼ� = RoundEx(objBalanceInfor.��ǰ���� - objBalanceInfor.�Ѹ��ϼ�, 6)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Function zlGetReadFeeDetailFromBalanceID(ByVal lng����ID As Long, int������Դ As Integer, bln���� As Boolean, ByVal blnNOMoved As Boolean, ByRef rsDetail_out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID,��ȡ������ϸ��Ϣ
    '���:lng����ID-����ID
    '     bln����-�Ƿ����ϼ�¼
    '     blnNOMoved-�Ƿ�����ת��
    '     int������Դ-1:����;2-סԺ;0 -�����סԺ
    '����:rsDetail_out-������ϸ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-02 19:14:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strFormat As String
    
    On Error GoTo errHandle
    strFormat = "99999999990." & String(IIf(gbytDec < 0, 1, gbytDec), "9")
    Select Case int������Դ
    Case 1  '����
        strSQL = "" & _
        "   (   Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,0 as ��ҳID,�վݷ�Ŀ,Ӥ����," & _
        "           Sum(���ʽ��) As ���ʽ��,����ʱ��,max(ҽ�����) as ҽ�����,Max(�Ƿ���) As �Ƿ���  " & vbCrLf & _
        "       From " & IIf(blnNOMoved, "H", "") & "������ü�¼ A " & vbCrLf & _
        "       where A.����ID=[1]  " & vbCrLf & _
        "       Group By ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,�վݷ�Ŀ,Ӥ����,����ʱ�� " & vbCrLf & _
        "    ) A "
        'strSQL = IIf(mblnNOMoved, "H", "") & "������ü�¼ A "
    Case 2  'סԺ
        strSQL = IIf(blnNOMoved, "H", "") & "סԺ���ü�¼ A"
    Case Else '�����סԺ
        strSQL = "" & _
        " (     Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,0 as ��ҳID,�վݷ�Ŀ,Ӥ����,���ʽ��,����ʱ��,ҽ�����,�Ƿ��� " & vbCrLf & _
        "       From " & IIf(blnNOMoved, "H", "") & "������ü�¼ A " & vbCrLf & _
        "       Where A.����ID=[1] " & vbCrLf & _
        "       Union ALL " & vbCrLf & _
        "       Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,��ҳID,�վݷ�Ŀ,Ӥ����,���ʽ��,����ʱ��,ҽ�����,�Ƿ��� " & vbCrLf & _
        "       From " & IIf(blnNOMoved, "H", "") & "סԺ���ü�¼ A " & vbCrLf & _
        "       Where A.����ID=[1]  " & vbCrLf & _
        " )  A"
    End Select
    
    strSQL = _
    "   Select Decode(�����־,1,'����',4,'����',Decode(Nvl(A.��ҳID,0),0,'','��'||Nvl(A.��ҳID,0)||'��')) As ����," & vbCrLf & _
    "         A.NO as ���ݺ�,Nvl(B.����,'δ֪') as ��������,decode(nvl(a.�Ƿ���,0),1,'***',Nvl(E.����,D.����)) as ��Ŀ," & vbCrLf & _
             IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "decode(nvl(a.�Ƿ���,0),1,'***',E1.����) as ��Ʒ��,", "") & vbCrLf & _
    "       A.�վݷ�Ŀ as ��Ŀ,Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����," & vbCrLf & _
    "       ltrim(rtrim(To_Char(" & IIf(bln����, "-1*", "") & "A.���ʽ��,'" & strFormat & "'))) as ���ʽ��," & vbCrLf & _
    "       To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��" & vbCrLf & _
    " From " & strSQL & ",���ű� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & vbCrLf & _
    " Where A.��������ID=B.ID And A.�շ�ϸĿID=D.ID" & vbCrLf & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & vbCrLf & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & vbCrLf, "") & _
    "       And A.����ID=[1]" & vbCrLf & _
    " Order by ���� Desc,����ʱ�� Desc,���ݺ� Desc,A.���"
    Set rsDetail_out = zlDatabase.OpenSQLRecord(strSQL, "���ݽ���ID��ȡ���������ϸ����", lng����ID)
    zlGetReadFeeDetailFromBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetExceptionBalanceData(ByVal bytType As Byte, ByRef dtStartDate As Date, _
    ByVal dtEndDate As Date, ByVal str����Ա As String, _
    rsErrData_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�쳣�Ľ�������
    '���:bytType:0-�쳣�Ľ��ʼ�¼;1-�쳣�Ľ������ϼ�¼
    '     bytDateRange-���ڷ�Χ:
    '����:rsErrData_Out-�����쳣�Ľ�������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-07 19:35:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere  As String, strTable As String, strSQL As String
    Dim rsTemp As ADODB.Recordset, str��Լ����IDs As String
    Dim cllFilter As Collection, cllPati As Collection
    Dim objPati As clsPatientInfo
    
    On Error GoTo errHandle
    strWhere = "  And A.�շ�ʱ�� Between [1] And [2] And A.����Ա���� = [3] And A.����״̬ = 1"
    If bytType = 0 Then
        strWhere = strWhere & " And A.��¼״̬ In (1,3)"
    Else
        strWhere = strWhere & "And A.��¼״̬ = 2"
    End If
    
    strTable = "" & _
    " Select A.ID ,1 as סԺ��־,0 as �����־,A.NO,A.ʵ��Ʊ��,A.����ID,A.��ҳID,A.��ʼ����,A.��������," & _
    "       Max(A.��¼״̬) As ��¼״̬,Sum(B.���ʽ��) As ���ʽ��,A.����Ա����,A.�շ�ʱ��," & _
    "       A.��;����,A.ԭ�� as ��Լ��λ,A.��������,Max(b.����ID) As ���ò���ID,Max(b.��ʶ��) As ��ʶ��, " & _
    "       Max(b.����) As ����,Max(b.�Ա�) As �Ա�,Max(b.����) As ����,Max(b.�ѱ�) As �ѱ�" & _
    " From ���˽��ʼ�¼ A,סԺ���ü�¼ B" & _
    " Where A.ID=B.����ID " & strWhere & _
    " Group By A.ID ,A.NO,A.ʵ��Ʊ��,A.����ID,a.��ҳid,A.��ʼ����,A.��������,A.����Ա����,A.�շ�ʱ��," & _
    "   A.��;����,A.ԭ��,A.�������� "
    
    strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & _
        Replace(Replace(strTable, "סԺ���ü�¼", "������ü�¼"), "1 as סԺ��־,0 as �����־", "0 as סԺ��־,1 as �����־")
    
    strSQL = "" & _
    " Select A.ID as ����ID,decode(סԺ��־,1,decode(�����־,1,3,2),1) As ��־," & _
    "        decode(A.��������,1,'�������',2,'סԺ����','') As �������� , " & _
    "        Decode(D.����,NULL,NULL,'��') as ҽ��,A.NO as ���ݺ�,A.ʵ��Ʊ�� As Ʊ�ݺ�," & _
    "        Decode(A.����ID,Null,' ',A.����ID) As ����ID," & _
    "        Decode(Nvl(A.��������,0),2,' ',Decode(A.����ID,Null,' ',a.��ʶ��)) As �����," & _
    "        Decode(Nvl(A.��������,0),1,' ',Decode(A.����ID,Null,' ',a.��ʶ��)) As סԺ��," & _
    "        Decode(A.����ID,Null,A.��Լ��λ,a.����) As ����," & _
    "        Decode(A.����ID,Null,' ',a.�Ա�) As �Ա�," & _
    "        Decode(A.����ID,Null,' ',a.����) As ����," & _
    "        Decode(A.����ID,Null,' ',a.�ѱ�) As �ѱ�," & _
    "        To_Char(A.��ʼ����,'YYYY-MM-DD') As ��ʼ����,To_Char(A.��������,'YYYY-MM-DD') As ��������," & _
    "        To_Char(Decode(A.��¼״̬,2,-1,1) *A.���ʽ��,'999999999" & gstrDec & "') as ���ʽ��," & _
    "        A.����Ա���� as ����Ա,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��," & _
    "        Decode(Nvl(A.��;����,0),1,'��',' ') ��;����,A.��¼״̬ as ��¼״̬,A.���ò���ID" & _
    " From ( " & strTable & ") A,���ս����¼ D,��Ա�� N" & _
    " Where  A.����Ա����=N.���� " & _
    "       And (N.վ��='" & gstrNodeNo & "' Or N.վ�� is Null) And A.id = D.��¼ID(+) and D.����(+)=2" & _
    " Order by �շ�ʱ�� Desc,���ݺ� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����쳣����", dtStartDate, dtEndDate, UserInfo.����)
    
    Do While Not rsTemp.EOF
        If Val(NVL(rsTemp!����ID)) = 0 And NVL(rsTemp!����) = "" Then
            If InStr("," & str��Լ����IDs & ",", "," & rsTemp!���ò���ID & ",") = 0 Then
                str��Լ����IDs = str��Լ����IDs & "," & rsTemp!���ò���ID
            End If
        End If
        rsTemp.MoveNext
    Loop
    
    If str��Լ����IDs = "" Then
        Set rsErrData_Out = rsTemp
    Else
        'ȡ��Լ��λ����
        str��Լ����IDs = Mid(str��Լ����IDs, 2)
        If gobjSquare.objOneCardComLib.zlGetMultiPatiInforFromPatiID(str��Լ����IDs, cllPati) = False Then Exit Function
        
        Set rsErrData_Out = zlDatabase.CopyNewRec(rsTemp)
        Do While Not rsErrData_Out.EOF
            If Val(NVL(rsErrData_Out!����ID)) = 0 And NVL(rsErrData_Out!����) = "" Then
                    Set objPati = cllPati("_" & rsTemp!���ò���ID)
                    rsErrData_Out!���� = objPati.������λ
            End If
            rsErrData_Out.MoveNext
        Loop
    End If
    If rsErrData_Out.RecordCount > 0 Then rsErrData_Out.MoveFirst
    
    zlGetExceptionBalanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCheck�������(lng����ID As Long, lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϲ����Ƿ������
    '���:lng����ID-����ID
    '     lng��ҳID-��ҳID
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-18 13:22:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPati As clsPatientInfo
    
    On Error GoTo ErrHandler
    Set objPati = New clsPatientInfo
    objPati.����ID = lng����ID
    objPati.��ҳID = lng��ҳID
    If zlGetPatiInfoByPage(objPati) = False Then Exit Function
    
    '49501
    If gTy_System_Para.byt������˷�ʽ = 0 Then
        zlCheck������� = (objPati.��˱�־ >= 1)
    Else
        zlCheck������� = (objPati.��˱�־ > 1)
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCheckPatiIsVerfy(ByVal bytEditType As gBalanceBill, ByVal objPati As clsPatientInfo, ByVal strPrivs As String, objBalanceAllCons As clsBalanceAllCon, _
    Optional ByRef strMessage As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡���Ƿ����
    '     objBalanceAllCons-��ǰ����
    '����:strMessage-������Ϣ
    '     objBalanceAllCons-����strUnAuditTime���Ե�סԺ����
    '����:���˺�
    '����:2015-01-05 14:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnAll As Boolean, lng��ҳID As Long, i As Long
    Dim varData As Variant
    On Error GoTo errHandle
    
    '���ﲻ���м��
    If bytEditType = g_Ed_������� Or objPati Is Nothing Then zlCheckPatiIsVerfy = True: Exit Function
    
    If InStr(strPrivs, ";δ��˲�����;����;") > 0 Or InStr(strPrivs, ";δ��˲��˳�Ժ����;") > 0 Then zlCheckPatiIsVerfy = True: Exit Function
    If objPati.��ҳID = 0 Then zlCheckPatiIsVerfy = True: Exit Function
    
    If CStr(objPati.��ҳID) = objBalanceAllCons.strAllTime Then  'ֻ�����һ��δ��
        If objPati.��˱�־ = 0 Then
            strMessage = "��ǰ����δ��ˣ��㲻�ܶ�δ��˵Ĳ��˽��н��ʡ�"
            Exit Function
        End If
        zlCheckPatiIsVerfy = True: Exit Function
    End If
    blnAll = True
    varData = Split(objBalanceAllCons.strAllTime, ",")
    For i = 0 To UBound(varData)
        lng��ҳID = Val(varData(i))
        If lng��ҳID <> 0 Then
            If Not zlCheck�������(objPati.����ID, lng��ҳID) Then
                 objBalanceAllCons.strUnAuditTime = objBalanceAllCons.strUnAuditTime & "," & lng��ҳID
            Else
                blnAll = False
            End If
        Else
            blnAll = False
        End If
    Next
    If objBalanceAllCons.strUnAuditTime <> "" Then objBalanceAllCons.strUnAuditTime = Mid(objBalanceAllCons.strUnAuditTime, 2)
    If blnAll Then
        strMessage = "�ò�������סԺ���ö�û����ˣ����ܽ��н��ʣ�"
        Exit Function
    End If
    zlCheckPatiIsVerfy = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCheckIsThirdSwapFromBalanceID(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ�������Ľ����Ƿ������������
    '���:lng����ID-����ID
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-20 16:09:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select 1  From ����Ԥ����¼ where ����ID=[1] and �����ID is not null and mod(��¼����,10)<>1 and rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ������������", lng����ID)
    zlCheckIsThirdSwapFromBalanceID = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCheckMulitInterfaceNumValied(ByVal vsBlance As VSFlexGrid, ByRef objCard As Card, objBalanceInfor As clsBalanceInfo, Optional blnԤ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ͬʱ�����������Ͻӿ�(��������)
    '����:�����������Ͻӿڵ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strIDs As String, strCardTypeIDs As String, strTemp As String, objItem As clsBalanceItem, objItems As clsBalanceItems
    Dim intMousePointer As Integer
    Dim intCount As Integer, i As Long, int���� As Integer, str���㷽ʽ As String
    Dim varData As Variant, strErrMsg As String

    On Error GoTo errHandle
    
    intMousePointer = Screen.MousePointer
     If objCard Is Nothing Then zlCheckMulitInterfaceNumValied = True: Exit Function

    If blnԤ�� Or objCard.�ӿ���� <= 0 Then zlCheckMulitInterfaceNumValied = True: Exit Function

    strErrMsg = ""
    
   'ҽ����һ���ӿ�
   If objBalanceInfor.objInsure.���� <> 0 And objBalanceInfor.�Ƿ񱣴���ʵ� Then intCount = intCount + 1: strErrMsg = strErrMsg & "ҽ������:" & Format(objBalanceInfor.ҽ��֧���ϼ�, gstrDec)
   
   strIDs = "": strCardTypeIDs = "," & objCard.�ӿ���� '�ѵ�ǰ���㷽ʽ�ų����ظ�ʹ�ý��㷽ʽ���м��
   With vsBlance
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            If zlGetBalanceItemFromBalanceGrid(vsBlance, i, objItem) Then
                strTemp = objItem.�����ID & "," & objItem.��������ID
                If InStr("34", int����) > 0 And InStr(strIDs & "|", "|" & strTemp & "|") = 0 And objItem.���ѿ� = False And objItem.�����ID > 0 Then
                    If zlGetBalanceItemsFromVsBalanceGrid(vsBlance, objItem, objItems) = False Then Exit Function
                    strIDs = strIDs & "|" & strTemp
                    If InStr(strCardTypeIDs & ",", "," & objItem.�����ID & ",") = 0 Then
                        strCardTypeIDs = strCardTypeIDs & "," & objItem.�����ID
                        intCount = intCount + 1
                        
                        If objItem.objCard Is Nothing Then
                            strErrMsg = strErrMsg & vbCrLf & objItem.���㷽ʽ & ":" & objItems.������
                        Else
                            strErrMsg = strErrMsg & vbCrLf & objItem.objCard.���� & ":" & objItems.������
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    If intCount >= 3 Then
        Screen.MousePointer = 0
        Call MsgBox("ע��:" & vbCrLf & "   ��ϵͳĿǰֻ֧���������½ӿ�,�����Ѿ��������½ӿڽ���:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    zlCheckMulitInterfaceNumValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
End Function


Public Function zlCheckDelBalanceIsValiedFromVsDeposit(ByRef vsDeposit As VSFlexGrid, ByVal objThirdSwap As clsThirdSwap, ByRef objBalanceInfor As clsBalanceInfo, ByRef objCurDelItem As clsBalanceItem, _
    ByRef objItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ�˿���Ϣ���󣬼���˿��Ƿ�Ϸ�
    '���:objCurDelItem-��ǰ�˿���
    '     objThirdSwap-�����ӿ�
    '����:objItems_Out-��ǰ�˿���Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-29 10:16:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, objCard As Card, objItemsTemp As clsBalanceItems, objItemsPt As clsBalanceItems
    Dim bln���ѿ� As Boolean, lngCardTypeID As Long, lngCurCardTypeID As Long
    Dim dblMoney As Double, dbl��Ԥ�� As Double, dblDelMoney As Double, blnSingleDel As Boolean
    Dim i As Long, j As Long, intMousePointer As Integer, strErrMsg As String, strExpend As String, strDefaultBalance As String
    Dim blnFind As Boolean, blnAdd As Boolean, blnDelCash As Boolean
    
    On Error GoTo errHandle

      
    If objCurDelItem Is Nothing Then Exit Function
    Set objCard = objCurDelItem.objCard
    
    If objCard Is Nothing Then Exit Function
    
    lngCurCardTypeID = objCurDelItem.�����ID
    
    
    dblDelMoney = RoundEx(Abs(objCurDelItem.������), 2)
    
    intMousePointer = Screen.MousePointer
    
    If objBalanceInfor.��Ԥ���ϼ� = 0 Then
        Screen.MousePointer = 0
        MsgBox "��ǰ��Ԥ�����˿����ʹ�á�" & objCard.���� & "�������˿������", vbInformation + vbOKOnly, gstrSysName
        Screen.MousePointer = intMousePointer
        Exit Function
    End If
    If dblDelMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox "��" & objCard.���� & "��δ�����˿�����ܽ����˿������", vbInformation + vbOKOnly, gstrSysName
        Screen.MousePointer = intMousePointer
        Exit Function
    End If
    
    blnSingleDel = objThirdSwap.zlThirdSwapIsSwapNOCall(lngCurCardTypeID, False, strErrMsg, strExpend)
    
    objCurDelItem.�Ƿ��˿�ֽ��� = blnSingleDel
    objCurDelItem.�Ƿ��˿� = True
    objCurDelItem.�Ƿ�Ԥ�� = True
    
    
    Set objItemsPt = New clsBalanceItems
    Set objItems_Out = New clsBalanceItems
    With vsDeposit
        For i = .Rows - 1 To 1 Step -1
            lngCardTypeID = Val(.TextMatrix(i, .ColIndex("�����ID")))
            bln���ѿ� = Val(.TextMatrix(i, .ColIndex("�Ƿ����ѿ�"))) = 1
            dbl��Ԥ�� = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
            
            If lngCurCardTypeID = lngCardTypeID And bln���ѿ� = False And Trim(.TextMatrix(i, .ColIndex("���ݺ�"))) <> "" And dbl��Ԥ�� > 0 Then
                If dblDelMoney = 0 Then Exit For
                If dblDelMoney >= dbl��Ԥ�� Then
                    dblMoney = dbl��Ԥ��
                    dblDelMoney = RoundEx(dblDelMoney - dbl��Ԥ��, 2)
                Else
                    dblMoney = dblDelMoney
                    dblDelMoney = 0
                End If
                
                Set objItem = zlCopyNewItemFromBalanceItem(objCurDelItem)
                If objCard Is Nothing Then Set objCard = zlGetCardFromCardType(lngCardTypeID, False, Trim(.TextMatrix(i, .ColIndex("���㷽ʽ"))))
                
                objItem.���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
                objItem.��������ID = Val(.TextMatrix(i, .ColIndex("��������ID")))
                objItem.�����ID = lngCardTypeID
                objItem.���� = Trim(.TextMatrix(i, .ColIndex("����")))
                objItem.������ˮ�� = Trim(.TextMatrix(i, .ColIndex("������ˮ��")))
                objItem.����˵�� = Trim(.TextMatrix(i, .ColIndex("����˵��")))
                objItem.������� = Trim(.TextMatrix(i, .ColIndex("�������")))
                objItem.������ = RoundEx(-1 * dblMoney, 6)
                objItem.����ժҪ = Trim(.TextMatrix(i, .ColIndex("ժҪ")))
                objItem.������� = IIf(objBalanceInfor.�������� = 1, True, False)
                objItem.�Ƿ��˿�ֽ��� = blnSingleDel
                objItem.����ID = objBalanceInfor.����ID
                objItem.����IDs = objBalanceInfor.����ID
                objItem.����ID = objBalanceInfor.����ID
                objItem.����ʱ�� = objBalanceInfor.����ʱ��
                objItem.�������� = objCard.��������
                objItem.�Ƿ�Ԥ�� = True
                objItem.�Ƿ��˿� = True
                objItem.�������� = 3 '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                objItem.Ԥ��ID = Val(.TextMatrix(i, .ColIndex("Ԥ��ID")))
                objItem.�Ƿ����� = objCard.�������Ĺ��� <> ""
                objItem.У�Ա�־ = 1
                
                Set objItem.objCard = objCard
                
                Set objItemsTemp = New clsBalanceItems
                objItemsTemp.AddItem objItem
                objItemsTemp.������ = objItem.������
                objItemsTemp.�շ����� = 1
                blnAdd = False
                
                 If Not objThirdSwap.zlThirdReturnCashCheck(objCard, objItemsTemp, blnDelCash, strDefaultBalance) Then
                    '1.��ֹ����
                    objItem.�Ƿ��������� = False
                    objItem.�Ƿ�ǿ������ = blnDelCash
                    objItem.�Ƿ�����ɾ�� = objItem.�Ƿ�ǿ������
                    blnAdd = True
                 Else
                    If blnDelCash = False Then  '�Ƿ�ȱʡ����
                          '�������֣�����ɾ��
                          objItem.�Ƿ�����༭ = False
                          objItem.�Ƿ�����ɾ�� = True
                          objItem.�Ƿ�ǿ������ = True
                          objItem.�Ƿ��������� = True: blnAdd = True
                      ElseIf strDefaultBalance <> "" Then
                          blnFind = False
                          For j = 1 To objItemsPt.Count
                              If objItemsPt(j).���㷽ʽ = strDefaultBalance Then
                                  objItemsPt(j).������ = objItemsPt(j).������ + objItem.������
                                  objItemsPt.������ = objItemsPt.������ + objItem.������
                                  blnFind = True
                                  Exit For
                              End If
                          Next
                          If Not blnFind Then
                              Set objItem = New clsBalanceItem
                              With objItem
                                  Set .objCard = zlGetCardFromBalanceName(strDefaultBalance)
                                  .���㷽ʽ = strDefaultBalance
                                  .������ = RoundEx(-1 * dblMoney, 6)
                                  .�Ƿ��˿� = True
                                  .�Ƿ�����༭ = False
                                  .�Ƿ�����ɾ�� = True
                                  .�������� = .objCard.��������
                                  .�������� = 0 '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                                  .Tag = "ָ��Ԥ���˿�"
                              End With
                              objItemsPt.AddItem objItem
                              objItemsPt.������ = RoundEx(objItemsPt.������ + objItem.������, 6)
                          End If
                      End If
                 End If
                 
                 If blnAdd Then  'δ�ҵ�����Ҫ����
                    objItems_Out.AddItem objItem
                    objItems_Out.������ = RoundEx(objItems_Out.������ + objItem.������, 2)
                 End If
            End If
        Next
    End With
    If objItems_Out.Count <> 0 And Not blnSingleDel Then
        '�ཻ��һ���˿�
        Set objItem = zlCopyNewItemFromBalanceItem(objCurDelItem)
        If objCard Is Nothing Then Set objCard = zlGetCardFromCardType(lngCardTypeID, False, objCurDelItem.���㷽ʽ)
        objItem.�Ƿ��˿� = True
        For i = 1 To objItems_Out.Count  'ֻҪһ��������������ģ������������ж�Ӧ�Ĵ���
            If objItems_Out(i).�Ƿ�����ɾ�� Then objItem.�Ƿ�����ɾ�� = True
            If objItems_Out(i).�Ƿ�ǿ������ Then objItem.�Ƿ�ǿ������ = True
            If objItems_Out(i).�Ƿ��������� Then objItem.�Ƿ��������� = True
        Next
        
        Set objItem.objTag = objItems_Out
        objItem.������ = objItems_Out.������
        Set objItems_Out = New clsBalanceItems
        objItems_Out.AddItem objItem
        objItems_Out.������ = objItems_Out.������ + objItem.������
    End If
        
    '������ͨ�Ľ��㷽ʽ
    For Each objItem In objItemsPt
        objItems_Out.AddItem objItem
        objItems_Out.������ = objItems_Out.������ + objItem.������
    Next
    
    If objItems_Out.������ = 0 Then
        Screen.MousePointer = 0
        MsgBox "��Ԥ������������У������ڡ�" & objCard.���� & "����Ԥ��������øý�����Ϣ�����˿������", vbInformation + vbOKOnly, gstrSysName
        Screen.MousePointer = intMousePointer
        Set objItems_Out = Nothing
        Exit Function
    End If
    
    If RoundEx(Abs(objItems_Out.������), 2) < RoundEx(Abs(objCurDelItem.������), 2) Then
        Screen.MousePointer = 0
        MsgBox "��Ԥ������������У���" & objCard.���� & "����ԭʼ������С���˱����˿�������øý�����Ϣ�����˿������" & vbCrLf & _
               "ԭʼ������:" & Format(objItems_Out.������, "0.00") & vbCrLf & _
               "�����˿���:" & Format(objCurDelItem.������, "0.00"), vbInformation + vbOKOnly, gstrSysName
        Screen.MousePointer = intMousePointer
        Set objItems_Out = Nothing
        Exit Function
    End If
    zlCheckDelBalanceIsValiedFromVsDeposit = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Screen.MousePointer = intMousePointer
End Function


Public Sub zlSetVsBalanceEditStatus(ByVal vsBlance As VSFlexGrid, ByVal objItem As clsBalanceItem, Optional blnSetRowData As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ñ༭״̬
    '���:blnSetRowData-�Ƿ�objItem���ø�Rowdata����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-09-03 10:06:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    If objItem Is Nothing Then Exit Sub
    
    lngRow = objItem.�к�
    With vsBlance
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        .TextMatrix(lngRow, .ColIndex("����״̬")) = IIf(objItem.�Ƿ����, 1, 0)
        .TextMatrix(lngRow, .ColIndex("�༭״̬")) = IIf(objItem.�Ƿ�����༭, 1, 0) & "|" & IIf(objItem.�Ƿ�����ɾ��, 1, 0)
        If blnSetRowData Then .RowData(lngRow) = objItem
    End With
End Sub




Public Function zlGetItemFromVsDepositRow(ByVal vsDeposit As VSFlexGrid, ByVal lngRow As Long, ByRef objItem_Out As clsBalanceItem) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ��Ԥ�����У���ȡ���е�Item����
    '���:vsDeposit-Ԥ�����б�
    '����:objItem_Out-��ȡָ����Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-06-14 15:14:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngCardTypeID As Long, bln���ѿ� As Boolean, dblMoney As Double
    
    Err = 0: On Error GoTo errHandle:
    Set objItem_Out = New clsBalanceItem
    '���������˿���Ϣ��ת����Ϣ
    With vsDeposit
        dblMoney = Val(.TextMatrix(lngRow, .ColIndex("��Ԥ��")))
        lngCardTypeID = Val(.TextMatrix(lngRow, .ColIndex("�����ID")))
        bln���ѿ� = Val(.TextMatrix(lngRow, .ColIndex("�Ƿ����ѿ�"))) = 1
        
        Set objItem_Out.objCard = zlGetCardFromCardType(lngCardTypeID, bln���ѿ�, Trim(.TextMatrix(lngRow, .ColIndex("���㷽ʽ"))))
        objItem_Out.�Ƿ�ת�� = Val(.TextMatrix(lngRow, .ColIndex("�Ƿ�ת�ʼ�����"))) = 1
        objItem_Out.�������� = Val(.TextMatrix(lngRow, .ColIndex("��������")))
        objItem_Out.����IDs = ""
        objItem_Out.������ˮ�� = Trim(.TextMatrix(lngRow, .ColIndex("������ˮ��")))
        objItem_Out.����˵�� = Trim(.TextMatrix(lngRow, .ColIndex("����˵��")))
        objItem_Out.���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
        objItem_Out.��������ID = Val(.TextMatrix(lngRow, .ColIndex("��������ID")))
        objItem_Out.������ = Val(.TextMatrix(lngRow, .ColIndex("��Ԥ��")))
        objItem_Out.���㷽ʽ = Trim(.TextMatrix(lngRow, .ColIndex("���㷽ʽ")))
        objItem_Out.������� = Trim(.TextMatrix(lngRow, .ColIndex("�������")))
        objItem_Out.����ժҪ = Trim(.TextMatrix(lngRow, .ColIndex("ժҪ")))
        If lngCardTypeID <> 0 Then
            objItem_Out.�������� = IIf(Not bln���ѿ�, 3, 5)    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        ElseIf objItem_Out.�������� = 7 Then
            objItem_Out.�������� = 4
        Else
            objItem_Out.�������� = 0 '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        End If
        objItem_Out.�Ƿ�Ԥ�� = True
        If .ColIndex("���") >= 0 Then
            objItem_Out.δ�˽�� = Val(.TextMatrix(lngRow, .ColIndex("���")))
        End If
        If .ColIndex("���") >= 0 Then objItem_Out.ԭʼ��� = Val(.TextMatrix(lngRow, .ColIndex("���")))
        objItem_Out.�����ID = lngCardTypeID
        objItem_Out.���ѿ� = bln���ѿ�
        objItem_Out.�Ƿ����� = IIf(objItem_Out.objCard.�������Ĺ��� <> "", True, False)
        objItem_Out.����ʱ�� = CDate(.TextMatrix(lngRow, .ColIndex("�տ�����")))
        objItem_Out.�Ƿ��������� = objItem_Out.objCard.�Ƿ�����
        objItem_Out.Ԥ��ID = Val(.TextMatrix(lngRow, .ColIndex("Ԥ��ID")))
        
        objItem_Out.���ѿ�ID = 0
        objItem_Out.�Ƿ��˿�ֽ��� = True
        objItem_Out.�Ƿ��˿� = False
        objItem_Out.�Ƿ�����༭ = False
        objItem_Out.�Ƿ�����ɾ�� = True
    End With
    zlGetItemFromVsDepositRow = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlReSetOppositePayMoneyFromItems(ByRef objCurItems As clsBalanceItems)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ŀ�Ľ�����ȡ����
    '���:objItems-��Ŀ��
    '����:objItems-���صķ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-09-05 12:02:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long

    On Error GoTo errHandle
    
    If objCurItems Is Nothing Then Exit Sub
    objCurItems.������ = RoundEx(-1 * objCurItems.������, 6)
    For i = 1 To objCurItems.Count
        Call zlReSetOppositePayMoneyFromItem(objCurItems(i))
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Sub zlReSetOppositePayMoneyFromItem(ByRef objCurItem As clsBalanceItem)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ŀ�Ľ�����ȡ����
    '���:objCurItem-��ǰ��Ŀ��
    '����:objCurItem-���صķ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-09-05 12:02:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItems As clsBalanceItems

    On Error GoTo errHandle
    If objCurItem Is Nothing Then Exit Sub
    
    objCurItem.������ = RoundEx(-1 * objCurItem.������, 6)
    Set objItems = objCurItem.objTag
    If objItems Is Nothing Then Exit Sub
    objItems.������ = RoundEx(-1 * objItems.������, 6)
    For i = 1 To objItems.Count
        objItems(i).������ = RoundEx(-1 * objItems(i).������, 6)
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlGetFromIDToBalanceData(ByVal lng����ID As Long, ByVal blnNOMoved As Boolean, _
    ByRef rsOutBalance As ADODB.Recordset, Optional blnView As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID����ȡ��������
    '���:lng����ID-����ID
    '     blnNoMoved-�Ƿ��Ѿ�ת�Ƶ��󱸱���
    '     blnView-�Ƿ����
    '����:rsOutBalance-��������
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 15:32:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, objCard As Card
    Dim rsNew  As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim blnExistOfflineYb As Boolean
    
    On Error GoTo errHandle
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�;6-����
    strSQL = "" & _
    "   Select  A.ID, " & _
    "        Case when Mod(A.��¼����,10)=1 then 1  " & _
    "             when (nvl(M.����,0)=3 or nvl(M.����,0)=4) and nvl(a.�����ID,0)=0  then 2 " & _
    "             when nvl(A.�����ID,0)<>0  then  3 " & _
    "             when nvl(M.����,0)=9 then 6 " & _
    "             else 0 end as ����, " & _
    "        Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,A.��Ԥ��,A.ժҪ,A.�����ID,A.���㿨���,A.�������,A.����,A.������ˮ��," & vbNewLine & _
    "        nvl(C1.���ƿ�,0) as ���ƿ�, nvl(C1.�Ƿ�����,0) as �Ƿ�����," & vbNewLine & _
    "        nvl(C1.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, 0 as �Ƿ�ת�ʼ�����," & vbNewLine & _
    "        Decode(C1.�Ƿ�����,NULL,0,1) as �Ƿ�����,C1.����  as ���������," & vbNewLine & _
    "        A.����˵��,A.�������,A.У�Ա�־,decode(nvl(M.����,0),3,1,4,1,0) as ҽ��,0 as ���ѿ�id," & _
    "        nvl(M.����,0) as ��������,A.ID as Ԥ��ID,A.��������ID,nvl(a.���ӱ�־,0) as ���ӱ�־" & vbNewLine & _
    "   From  ����Ԥ����¼ A ,���㷽ʽ M,���ѿ����Ŀ¼ C1" & _
    "   Where A.����ID= [1] And A.���㷽ʽ=M.����(+) And a.���㿨��� = c1.���(+)  " & _
    "         And ( nvl(A.���㿨���,0)=0 OR  Mod(a.��¼����,10) =1) "
    
    If Not blnView Then
        '--�������תסԺʱ��һ�ν��㣨ҽ�ƿ�֧�������������ͨ�����ת�������˶��סԺԤ�����ݣ���ЩԤ�����ݵĹ�������ID��ͬ
        '--�ڽ���Ԥ���˿�ʱ����������ID��ͬ�ļ�¼��Ԥ����¼��ֻ��һ���������˿���Ϣ���ж��������ԣ�
        '   ���ų�Ԥ����¼�е����ݣ����������˿���Ϣ�еĽ����㷽ʽΪNULL���������˿���Ϣ�����浥������
        '--����Ԥ���˿�ʱ���Ͳ����ڷ�Ԥ������˿���Բ���Ԥ����¼�������˿���Ϣ�Ĺ�����������Ϊ(����id,�����id��
        strSQL = strSQL & _
        "         And Not Exists (Select 1 From �����˿���Ϣ" & _
        "                         Where ����id = a.����id And �����id = a.�����id And a.��¼���� = 2 And a.��Ԥ�� < 0)"

        strSQL = strSQL & " Union ALL " & _
        " Select a.Id, 0 As ����, Mod(a.��¼����, 10) As ��¼����, '' As ���㷽ʽ, a.��Ԥ��, '' As ժҪ," & _
        "       Null As �����id, Null As ���㿨���, '' As �������, '' As ����, '' As ������ˮ��, " & _
        "       0 As ���ƿ�, 0 As �Ƿ�����, 0 As �Ƿ�ȫ��, 0 As �Ƿ�ת�ʼ�����, 0 As �Ƿ�����, '' As ���������," & _
        "        '' As ����˵��, a.�������, 0 As У�Ա�־, 0 As ҽ��, 0 As ���ѿ�id, 0 As ��������," & _
        "       a.Id As Ԥ��id, Null As ��������id, 0 As ���ӱ�־" & _
        " From ����Ԥ����¼ A" & _
        " Where a.����id = [1] And Mod(a.��¼����, 10) <> 1 And a.�����id Is Not Null" & _
        "       And Exists (Select 1 From �����˿���Ϣ" & _
        "                   Where ����id = a.����id And �����id = a.�����id And a.��¼���� = 2 And a.��Ԥ�� < 0)"
    End If
    
    strSQL = strSQL & " Union ALL " & _
    "   Select A.ID,5 as  ����,Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,-1*nvl(b.Ӧ�ս��,0) as ��Ԥ��,A.ժҪ,A.�����ID,A.���㿨���," & _
    "        A.�������,B.����,B.������ˮ��,nvl( M.���ƿ�,0) as ���ƿ�, " & _
    "        nvl( M.�Ƿ�����,0) as �Ƿ�����,nvl(M.�Ƿ�ȫ��,0) as �Ƿ�ȫ��,0 as �Ƿ�ת�ʼ�����," & _
    "        nvl(M.�Ƿ�����,0) as  �Ƿ�����," & _
    "        M.���� as ���������,A.����˵��,A.�������,A.У�Ա�־,0 as ҽ��,B.���ѿ�id,M1.���� as ��������,A.ID as Ԥ��ID,A.��������ID,nvl(a.���ӱ�־,0) as ���ӱ�־ " & _
    "   From ����Ԥ����¼ A ,���˿������¼ B,���ѿ����Ŀ¼ M,���㷽ʽ M1 " & _
    "   Where  a.Id = b.����Id  and a.���㿨��� = m.��� And A.���㷽ʽ=M1.����(+) And A.����ID = [1] and Mod(A.��¼����,10)<>1 "
       
    strSQL = "" & _
    " Select a.����, a.��¼����, a.���㷽ʽ, a.ժҪ, a.�����id, a.���������, a.���ƿ�, a.���㿨���," & _
    "        a.�������, a.����, a.������ˮ��, A. ����˵��, a.�������, a.У�Ա�־, a.ҽ��, a.���ѿ�id," & _
    "        a.�Ƿ�����, a.�Ƿ�ȫ��, a.�Ƿ�ת�ʼ�����, a.�Ƿ�����, Nvl(a.��Ԥ��, 0) As ��Ԥ��," & _
    "        a.�������� As ����, a.Ԥ��id, a.��������id, a.���ӱ�־" & _
    " From (" & strSQL & ") A" & _
    " Order by ����"
    
    If blnNOMoved Then
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
        strSQL = Replace(strSQL, "���˿������¼", "H���˿������¼")
        strSQL = Replace(strSQL, "�����˿���Ϣ", "H�����˿���Ϣ")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������", lng����ID)
    blnExistOfflineYb = False
    rsTemp.Filter = "�����ID<>0"
    If rsTemp.RecordCount > 0 Then
        rsTemp.Filter = ""
        Set rsNew = zlDatabase.CopyNewRec(rsTemp)
        
        rsNew.Filter = "�����ID<>0"
        Do While Not rsNew.EOF
            If ZlGetPayCard(Val(NVL(rsNew!�����ID)), objCard) Then
                rsNew!��������� = objCard.����
                rsNew!���ƿ� = IIf(objCard.���ƿ�, 1, 0)
                rsNew!�Ƿ����� = IIf(objCard.�������Ĺ��� = "", 0, 1)
                rsNew!�Ƿ�ȫ�� = IIf(objCard.�Ƿ�ȫ��, 1, 0)
                rsNew!�Ƿ����� = IIf(objCard.�Ƿ�����, 1, 0)
                rsNew!�Ƿ�ת�ʼ����� = IIf(objCard.�Ƿ�ת�ʼ�����, 1, 0)
                rsNew.Update
            End If
            rsNew.MoveNext
        Loop
        rsNew.Filter = ""
    Else
        rsTemp.Filter = ""
        Set rsNew = rsTemp
    End If
 
    
    Set rsOutBalance = rsNew
    
    
    zlGetFromIDToBalanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetOldOfflineBalanceFromBlanaceID(ByVal lng����ID As Long, ByVal objBalanceInfor As clsBalanceInfo, _
    ByVal strOffLineBalances As String, ByRef objBalanceItems As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԭʼ���ѻ�ҽ�����㼯
    '���:lng����id
    '     strOffLineBalances-�ѻ����㷽ʽ,��ʽ�����㷽ʽ|���㷽ʽ...
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-12-23 19:32:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL  As String
    Dim objItem As clsBalanceItem
    Dim objCard As Card
    
    On Error GoTo errHandle
    
    Set objBalanceItems = New clsBalanceItems
    If strOffLineBalances = "" Then Exit Function
    
    strSQL = "" & _
    "   Select ���㷽ʽ,mod(��¼����,10) as ��¼����,��Ԥ�� From ����Ԥ����¼  " & _
    "   Where ����ID IN(Select distinct B.ID From ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
    "                   Where a.ID=[1] And  A.NO=B.NO And B.��¼״̬ in (1,3)) And Mod(��¼����,10)<>1 And instr([2] ,'|'|| ���㷽ʽ||'|')>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡԭʼ���ʽ��-�ѻ�ҽ��", lng����ID, "|" & strOffLineBalances & "|")
    Do While Not rsTemp.EOF
        Set objItem = New clsBalanceItem
        Set objCard = zlGetCardFromCardType(0, False, NVL(rsTemp!���㷽ʽ))
       
        objItem.���㷽ʽ = NVL(rsTemp!���㷽ʽ)
        objItem.����ID = objBalanceInfor.����ID
        objItem.����IDs = objBalanceInfor.����ID
        objItem.����ID = objBalanceInfor.����ID
        objItem.����ʱ�� = objBalanceInfor.����ʱ��
        objItem.�������� = 0
        objItem.������ = RoundEx(rsTemp!��Ԥ��, 6)
        objItem.�Ƿ�����ɾ�� = True
        objItem.�Ƿ��������� = True
        objItem.�Ƿ�����༭ = False
        objBalanceItems.AddItem objItem
        
        rsTemp.MoveNext
    Loop
    zlGetOldOfflineBalanceFromBlanaceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If


End Function
Public Function zlGetLedDisplayBankDatasFromVsBalance(ByVal vsBlance As VSFlexGrid, ByRef cllBanks_out As Collection, ByVal dblҽ���ʻ���� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ����б���ȡ��ʾ��Led�ϵĽ������ݼ�
    '���:objPati-������Ϣ��
    '
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-09-26 16:21:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllYBBanks As Collection, cllThirdBanks As Collection, cllOldCardOneBanks As Collection
    Dim cllPTBanks As Collection
    Dim i As Long
    
    On Error GoTo errHandle
    
    Set cllYBBanks = New Collection
    Set cllThirdBanks = New Collection
    Set cllOldCardOneBanks = New Collection
    Set cllPTBanks = New Collection
    
    With vsBlance
        For i = 1 To .Rows - 1
            'ҽ������
            If .TextMatrix(i, .ColIndex("���㷽ʽ")) <> "" Then
                '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                Select Case Val(.TextMatrix(i, .ColIndex("����")))
                Case 2 'ҽ��
                    cllYBBanks.Add Array(.TextMatrix(i, .ColIndex("���㷽ʽ")) & ":", Format(Val(.TextMatrix(i, .ColIndex("������"))), "0.00"))
                Case 3 '�����ӿڽ���
                    cllThirdBanks.Add Array(.TextMatrix(i, .ColIndex("���㷽ʽ")) & ":", Format(Val(.TextMatrix(i, .ColIndex("������"))), "0.00"))
                Case 4 ' һ��ͨ����
                    cllOldCardOneBanks.Add Array(.TextMatrix(i, .ColIndex("���㷽ʽ")) & ":", Format(Val(.TextMatrix(i, .ColIndex("������"))), "0.00"))
                Case Else
                    cllPTBanks.Add Array(.TextMatrix(i, .ColIndex("���㷽ʽ")) & ":", Format(Val(.TextMatrix(i, .ColIndex("������"))), "0.00"))
                End Select
            End If
        Next
    End With
    
    Set cllBanks_out = New Collection
    
    If cllYBBanks.Count <> 0 Then
        cllBanks_out.Add Array("ҽ������:", Format(dblҽ���ʻ����, "0.00"))
        For i = 1 To cllYBBanks.Count
            cllBanks_out.Add cllYBBanks(i)
        Next
    End If
    
    Set cllYBBanks = Nothing
    
    If cllThirdBanks.Count <> 0 Then
        cllBanks_out.Add Array("һ��ͨ����:", "")
        For i = 1 To cllThirdBanks.Count
            cllBanks_out.Add cllThirdBanks(i)
        Next
    End If
    Set cllThirdBanks = Nothing
    
    If cllOldCardOneBanks.Count <> 0 Then
        cllBanks_out.Add Array("һ��ͨ����(��):", "")
        For i = 1 To cllOldCardOneBanks.Count
            cllBanks_out.Add cllThirdBanks(i)
        Next
    End If
    Set cllOldCardOneBanks = Nothing
    
    If cllPTBanks.Count <> 0 Then
        For i = 1 To cllPTBanks.Count
            cllBanks_out.Add cllPTBanks(i)
        Next
    End If
    Set cllPTBanks = Nothing
    zlGetLedDisplayBankDatasFromVsBalance = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetDiagIDFromComboxDiag(ByVal intComboxIdex As Integer, ByRef cboDiag As ComboBox) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰѡ�����ϻ�ȡ��ȡ���ID
    '���:intComboxIdex-����
    '     cboDiag-���������
    '����:ҽ��ID
    '����:���˺�
    '����:2019-01-24 11:22:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strҽ��ID As String
    On Error GoTo errHandle
    If intComboxIdex < 0 Or intComboxIdex > cboDiag.ListCount - 1 Or cboDiag.Tag = "" Then Exit Function
    zlGetDiagIDFromComboxDiag = Split(cboDiag.Tag & ",,,", ",")(intComboxIdex)
    Exit Function
errHandle:
    zlGetDiagIDFromComboxDiag = ""
End Function
Public Sub zlLoadDiagnosDataToCombox(ByVal frmMain As Object, ByVal rsAllDiagnos As ADODB.Recordset, ByRef cboDiag As ComboBox)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ϣ��Combox�ؼ�
    '���:str���IDs-ҽ��IDs
    '     cboDiag-������ϵ������ؼ�
    '     rsAllDiagnos-��ϵļ�¼��
    '����:���˺�
    '����:2019-01-23 18:03:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���ID As String, strTemp As String
    Dim lngWidth As Single, lngTemp As Single
    Dim j As Long
    On Error GoTo errHandle
    
  
    str���ID = zlGetDiagIDFromComboxDiag(cboDiag.ListIndex, cboDiag)
    cboDiag.Clear
    cboDiag.AddItem "�������"
    cboDiag.ListIndex = cboDiag.NewIndex
    cboDiag.Tag = "0"
    
    If rsAllDiagnos Is Nothing Then Exit Sub
    If rsAllDiagnos.State <> 1 Then Exit Sub
    lngWidth = cboDiag.Width
    
    With rsAllDiagnos
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            strTemp = zlFormatID(!���ID)
            If InStr("," & cboDiag.Tag & ",", "," & strTemp & ",") = 0 Then
                cboDiag.Tag = cboDiag.Tag & "," & strTemp
                cboDiag.AddItem NVL(!�������)
                
                j = frmMain.TextWidth("L") + 15
                If j * zlCommFun.ActualLen(NVL(!�������)) > 6465 Then
                    lngTemp = 6465
                Else
                    lngTemp = j * zlCommFun.ActualLen(NVL(!�������)) + frmMain.TextWidth("��") * 3
                End If
                If lngWidth < lngTemp Then lngWidth = lngTemp
                If strTemp = str���ID Then cboDiag.ListIndex = cboDiag.NewIndex
            End If
            .MoveNext
        Loop
    End With
    If lngWidth > cboDiag.Width Then
        Call zlcontrol.CboSetWidth(cboDiag.hWnd, lngWidth)
    Else
        Call zlcontrol.CboSetWidth(cboDiag.hWnd, cboDiag.Width)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlGetSelfPaymentMode(ByVal lngModule As Long, ByVal lng����ID As Long, ByVal str��ҳIDS As String, _
    ByVal lng����ID As Long, ByVal rsFeelists As ADODB.Recordset, _
    ByRef strSelfPaymentMode_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�Էѷ�ʽ
    '���:lngModule-����ģ���
    '     lng����ID-����ID
    '     str��ҳIds-��ǰ��ҳIDs,����ö��ŷ���
    '     lng����ID-�����ҽ�����ˣ���ҽ������󣬵��ñ��ӿڣ����뵱ǰ�Ľ���ID,
    '               �������ͨ���ˣ��ڶ�ȡ������ϸ�󣬵��ñ��ӿڣ�����IDΪ0,��������rsFeeLists��¼��
    '     rsFeeLists-����ı��ν��ʵķ�����ϸ���ݣ�סԺ,���ݺ� ,��Ŀ,��Ŀ,ID,���,��¼����,��¼״̬,ִ��״̬,��ҳID,���㵥λ,
    '    ����, �۸�,Ӧ�ս��,ʵ�ս��,δ����,���ʽ��, ͳ����,����, �շ����,�շ������,�ѱ�, Ӥ����, ִ�в���id,
    '    ����,��������ID,������, ���մ���id,�շ�ϸĿID,�����־,����, ҽ�����,ʱ��,�Ǽ�ʱ�䣩(�������ֽ��ʳ����еķ�����ϸ����
    '
    '����:strSelfPaymentMode_Out-���ؽ��㷽ʽ,���ӿڷ���trueʱ��Ч,��ʽΪ:���㷽ʽ1,������1,�������1||���㷽ʽ2,������2,�������2||....
    '����:��ȡ����true,���򷵻�False
    '����:���˺�
    '����:2019-05-05 14:41:27
    '����ʱ��:
    '     1.��ͨ����:�ڶ�ȡ���˷�����ϸ����ñ��ӿ�
    '     2.ҽ�����ˣ�����ҽ������󣬵��ñ��ӿ�
    'Ӧ�ó��������ƾ���
    '    1�����ƾ����û����財���û�ʿ������ݵĺ˶ԣ������ڽ��ʵ�ʱ�򲻻�����˵�Լ��Ǵ��ƾ����Ķ��������ý���Ա�ֹ�¼��Ļ�������©���ⲿ�֡�
    '    2�����ƾ�������㷨�ǣ������ҽ�����ˣ���Ҫ��ҽ���������֮��Ű����Էѽ���xx%����ģ���ͨ������ֱ�����Էѷ��õ�xx%��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetSelfPaymentMode = False
End Function

Public Function zlAddfinancialTrancsToBalanceList(ByRef vsBlance As VSFlexGrid, ByVal dblMoney As Double, Optional dblTranMoney_out As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ת�ʽ�����Ϣ
    '���:vsBlance-�����б�
    '     dblMoney-�����˿���
    '����:dblTranMoney_out-ת�˽��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-09-09 10:57:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str������� As String, str����ժҪ As String, str���� As String
    
    On Error GoTo errHandle
    dblTranMoney_out = 0
    If Not (gTy_System_Para.TY_Balance.blnԤ����ָ�����㷽ʽ And gTy_System_Para.TY_Balance.strԤ���˿���㷽ʽ <> "") Or dblMoney >= 0 Then
        zlAddfinancialTrancsToBalanceList = True: Exit Function
    End If
   
    Call ClearBalanceList(vsBlance, gTy_System_Para.TY_Balance.strԤ���˿���㷽ʽ)
    With vsBlance
        If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) <> "" Or .Rows <= 1 Then .Rows = .Rows + 1
        i = .Rows - 1
        
        Call zlPlugin_GetfinancialTrancsBalanceInfor(gTy_System_Para.TY_Balance.strԤ���˿���㷽ʽ, dblMoney, str�������, str����ժҪ, str����)
        
        .RowData(.Rows - 1) = ""
        .TextMatrix(i, .ColIndex("����")) = 0    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        .TextMatrix(i, .ColIndex("�����ID")) = 0
        .TextMatrix(i, .ColIndex("���ѿ�ID")) = 0
        .TextMatrix(i, .ColIndex("��������")) = 2    ''1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
        .TextMatrix(i, .ColIndex("�༭״̬")) = 0   ''0-��ֹɾ��;1-����༭���;2-������ɾ��;3-����ɾ�����޸Ľ��,4-��ֹɾ���ҽ�ֹ�޸ĵȵ�
        .TextMatrix(i, .ColIndex("����״̬")) = 0  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
        .TextMatrix(i, .ColIndex("�Ƿ�����")) = 0
        .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = 0
        .TextMatrix(i, .ColIndex("У�Ա�־")) = 0
        .TextMatrix(i, .ColIndex("�Ƿ�ת��")) = 0
        .TextMatrix(i, .ColIndex("�Ƿ�����")) = 0
        .TextMatrix(i, .ColIndex("���������")) = ""
        .TextMatrix(i, .ColIndex("���㷽ʽ")) = gTy_System_Para.TY_Balance.strԤ���˿���㷽ʽ
        .TextMatrix(i, .ColIndex("������")) = Format(dblMoney, "0.00")
        .TextMatrix(i, .ColIndex("�������")) = str�������
        .TextMatrix(i, .ColIndex("��ע")) = str����ժҪ
        .TextMatrix(i, .ColIndex("������ˮ��")) = ""
        .TextMatrix(i, .ColIndex("����˵��")) = ""
        .TextMatrix(i, .ColIndex("����")) = str����
        .Cell(flexcpData, i, .ColIndex("����")) = str����
        dblTranMoney_out = Format(dblMoney, "0.00")
    End With
    zlAddfinancialTrancsToBalanceList = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetDepoistRowFromBalanceList(ByVal vsBlance As VSFlexGrid, ByVal str���㷽ʽ As String) As Integer
    '���ܣ��ӽ����б��л�ȡ��ʣ��Ԥ���ĳ�ֽ��㷽ʽת�ˡ���
    '��Σ�vsBlance-����Ľ�����Ϣ�б�
    '      ��str���㷽ʽ-����ʣ����ΪԤ����Ľ��㷽ʽ
    Dim i As Integer
    
    If vsBlance.Rows <= 1 Then Exit Function
    With vsBlance
        If .ColIndex("���㷽ʽ") = -1 Or .ColIndex("��ע") = -1 Then Exit Function
        For i = 1 To vsBlance.Rows - 1
            If .TextMatrix(i, .ColIndex("���㷽ʽ")) = str���㷽ʽ And Val(.TextMatrix(i, .ColIndex("����"))) = 0 Then
                GetDepoistRowFromBalanceList = i: Exit For
            End If
        Next
    End With
End Function

Public Sub ClearBalanceList(ByVal vsBlance As VSFlexGrid, ByVal str���㷽ʽ As String)
    '���ܣ�Ԥ������Խ���ʱ�����������ʣ����Ԥ������
    '��Σ�vsBlance-����Ľ�����Ϣ�б�
    '      ��str���㷽ʽ-����ʣ����ΪԤ����Ľ��㷽ʽ
    Dim i As Integer
    
    If vsBlance.Rows <= 1 Then Exit Sub
    i = GetDepoistRowFromBalanceList(vsBlance, str���㷽ʽ)
    If i > 0 Then vsBlance.RemoveItem i

End Sub

Public Function zlCheckOtherSessionDoing(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ�����Ƿ����ڱ������Ự����
    '���:lng����ID-ָ���Ľ������
    '����:
    '����:�������Ựվ�÷���true,���򷵻�False
    '˵����"����Ԥ����¼.�Ự��"��ʽ��V$session.SID+'_'+V$session.SERIAL#
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If lng����ID = 0 Then zlCheckOtherSessionDoing = False: Exit Function
    
    strSQL = "Select 1" & vbNewLine & _
            " From ����Ԥ����¼ A, V$session B" & vbNewLine & _
            " Where a.�Ự�� = b.Sid || '_' || b.Serial# And a.����ID = [1] " & vbNewLine & _
            "       And b.Username Is Not Null And b.Audsid <> Userenv('sessionid')" & vbNewLine & _
            "       And Upper(b.Status) In ('ACTIVE', 'INACTIVE') And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ�����Ƿ����ڱ������Ự����", lng����ID)
    zlCheckOtherSessionDoing = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckBalanceOverFromBalanceID(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID,��鵱ǰ�����Ƿ��Ѿ����
    '����:������ɷ���true,���򷵻�False
    '����:���˺�
    '����:2019-09-17 10:53:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "Select 1 From ���˽��ʼ�¼ where id=[1] And nvl(����״̬,0)=0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ����״̬", lng����ID)
    zlCheckBalanceOverFromBalanceID = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlErrBalanceCheckFromPatiID(ByVal lng����ID As Long, ByRef strErrNo_Out As String, ByRef blnDel_Out As Boolean, _
    Optional ByVal strCheckNO As String, Optional ByVal str��ǰ���� As String = "����") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID��ǰ�Ľ���NO���ж��Ƿ�����쳣�Ľ��ʻ��˷ѵ���
    '���:lng����ID-����ID
    '     strCheckNo-��Ϊ��ʱ�������ʵ��ݺŽ��м��
    '����:
    '    strErrNo_Out-��������trueʱ����ʾ���ص��쳣����NO,����Falseʱ��������Ϊ��
    '    blnDel_Out-��ǰ���ص��쵥���Ƿ�Ϊ�쳣���Ͻ��ʵ���
    '����:0-�������쳣����
    '     1-�����쳣���ݣ���ѡ��Ϊ��������
    '     2-�����쳣���ݣ���ǰѡ��Ϊ��ֹ����
    '     3-�����쳣���ݣ���ǰѡ������쳣���ݽ����ؽ������
    '����:���˺�
    '����:2019-09-17 10:53:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lng����ID As Long, str����Ա���� As String, strTittle As String
 
     
    strErrNo_Out = "": blnDel_Out = False
     
    strSQL = " " & _
    "    Select  a.No, a.ID, a.����Ա����, decode(��¼״̬,2,2,1) As �쳣����,A.�շ�ʱ�� " & _
    "    From ���˽��ʼ�¼ A" & _
    "    Where nvl(����״̬,0) = 1" & IIf(strCheckNO = "", " And ����ID=[1]", " And No=[2]")
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����쳣�Ľ��ʵ���", lng����ID, strCheckNO)
    
    If rsTemp.EOF Then zlErrBalanceCheckFromPatiID = 0: Exit Function   '0-�������쳣����
    
    With rsTemp
        '�����ϴ��ڶ���쳣����Ҫ�ظ����
        Do While Not .EOF
            If zlCheckOtherSessionDoing(Val(NVL(rsTemp!ID))) Then
                MsgBox "ע��:" & vbCrLf & _
                "    �ò��˴����쳣��" & IIf(Val(NVL(rsTemp!�쳣����)) = 2, "����", "����") & "����(" & NVL(rsTemp!NO) & ")�������������ʴ��ڽ��в��� ,�����ڲ��ܽ���" & str��ǰ���� & "����!", vbInformation + vbOKOnly, gstrSysName
                zlErrBalanceCheckFromPatiID = 2
                Exit Function
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    

    strErrNo_Out = NVL(rsTemp!NO): lng����ID = Val(NVL(rsTemp!ID))
    blnDel_Out = Val(NVL(rsTemp!�쳣����)) = 2
    strTittle = IIf(Not blnDel_Out, "����", "����")
    str����Ա���� = NVL(rsTemp!����Ա����)
    
    
    If str����Ա���� <> UserInfo.���� Then
        '100703
         If MsgBox("ע��:" & vbCrLf & _
                    "    �ò��˴����쳣��" & strTittle & "����" & IIf(str����Ա���� <> UserInfo.����, ",�õ����ǲ���Ա[" & str����Ա���� & "]��ȡ��," & vbCrLf, "") & " ,���Ƿ��������" & str��ǰ���� & "����?" & vbCrLf & vbCrLf & _
                    "���ǡ��������쳣���ݽ��д���,��������" & str��ǰ���� & "����. " & vbCrLf & _
                    "���񡻴�����ֹ���ʲ���.", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            zlErrBalanceCheckFromPatiID = 1
         Else
           zlErrBalanceCheckFromPatiID = 2
         End If
        Exit Function
    End If
    If MsgBox("ע��:" & vbCrLf & _
                        "       �ò��˴����쳣��" & strTittle & "����(" & strErrNo_Out & "),���Ƿ���Ҫ���¶Ըõ��ݽ���" & strTittle & "?" & vbCrLf & vbCrLf & _
                        "���ǡ��������¶��쳣���ݽ���" & strTittle & vbCrLf & _
                        "���񡻴������쳣���ݽ��д���,��������" & str��ǰ���� & "����.", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        zlErrBalanceCheckFromPatiID = 1
    Else
        zlErrBalanceCheckFromPatiID = 3
    End If
End Function

Public Function zlGetBalanceIDFromBalanceNo(strNO As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���NO,��ȡԭʼ�Ľ���ID
    '���:strNo-����ID
    '����:����ԭʼ�Ľ���ID
    '����:���˺�
    '����:2020-03-25 18:52:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select ID From ���˽��ʼ�¼ Where NO=[1] And ��¼״̬ in (1,3)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݽ��ʵ���ȡԭ����ID", strNO)
    If rsTemp.EOF Then Exit Function
     zlGetBalanceIDFromBalanceNo = Val(NVL(rsTemp!ID))
End Function


