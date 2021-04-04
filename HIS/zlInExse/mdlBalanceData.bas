Attribute VB_Name = "mdlBalanceData"
Option Explicit
'*********************************************************************************************************************************************
'功能:获取相关结帐信息数据
'函数:
'  01：公共函数
'    0101. zlGetBalanceDataErrFromBalanceID:根据结帐ID获取相关的异常结帐数据信息
'    0102. zlGetBalanceDataFromBalanceID:根据结帐ID获取相关的结帐数据信息
'    0103. zlGetSaveThreeDelSwapBatchSQL:根据原交易信息(clsBalanceItem)批量获取保存三方退款信息的SQL
'    0104. zlGetBalanceItemSQLFromBalanceItem:根据结算对象(clsBalanceItem)获取结帐数据修正的SQL
'    0105. zlGetSaveThirdSwapDelSQLFromBalanceItem:根据结算对((clsBalanceItem))象获取三方交易退款保存SQL
'    0106. zlThirdDelSwapIsExsistFromBalanceID:根据结帐ID判断三方退款交易是否存在
'    0107. zlGet结算方式:获取结算方式(含应用场合)
'    0108. zlGetClassMoney：获取按收费类别汇总的记录集
'    0109. zlGetRemainderMoneyToPati：获取病人余额信息给病人对象
'    0110. zlGetDefaultHospitalizedDate：根据病人ID获取上次中途结帐时间
'    0111. zlIsCheck病历已接收:检查病历是否接收
'    0112. zlGetThirdMoneyInforRecordFromSwapID:根据交易ID,获取相关交易的结算信息集(含原始金额，已退金额，未退金额等)
'    0113. zlCheck病人审核:检查病人是否审核
'    0114. CheckPatiIsVerfy:检查指定病人是否已经审核
'    0115. zlComparePatiNumsIsDiff:比较两个住院次数是否一致
'    0116. zlCopyNewFeeData:根据业务类型的记录集，拷贝新的数据集
'    0117. zlCheckNoSettlementMoney:检查门诊留观病人是否存在未结费用金额
'    0118. zlErrBalanceCheckFromPatiID:根据病人ID或结帐单据号判断异常单据
'    0119. zlCheckBalanceOverFromBalanceID:根据结帐ID，判断是否当前结帐是否已经结帐成功
'    0120. zlCheckOtherSessionDoing:根据结帐ID，检查当前结帐是否被其他会话站用
'  02：一卡通接口相关
'    0201. zlGetCardFromBalanceName:根据结算方式名称，获取卡对象
'    0202. zlGetCardFromCardType:根据卡类别ID获取卡对象
'    0203  zlGetBalanceItemFromCardObject:根据卡对象，获取新的结算信息对象
'  03：结算列表相关函数
'    0300. zlInitBalanceGrid:初始化结算列表信息
'    0301. zlGetBalanceItemsFromVsBalanceGrid:根据结算网格及关联交易ID，获取指定结算数据集
'    0302. zlGetBalanceItemsFromRecord:根据结帐记录数据返回指定的结算数据集
'    0303. zlClearBalanceFromBalanceGrid:根据结算方式来清除结算信息行
'    0304. zlCheckBalancesIsExistFromCardTypeID:根据卡类别ID检查是否在结算列表中存在该结算类别的结算信息
'    0304. zlCheckVsBalanceIsExsitsFromCardObject:根据卡对象，检查结算列靓啊中是否存在指定的结算方式
'    0305. zlGetBalanceItemFromBalanceGrid:根据指定行获取结算项信息
'    0306. zlGetBalanceItemsFromCardObject:根据卡对象,从结算列表中获取相关的结算数据
'    0307. zlGetCancelBalancesFromVsBalanceGrid:根据结算列表，获取作废的结算信息
'    0308. zlAddBalanceDataToGridFromBalanceItems:根据结算信息对象，加载到结算列表中
'    0309  zlLoadBalanceItemsToVsGrid:根据结算信息对象，回载到结算列表
'    0310. zlGetBalanceNULLRow:获取空行
'    0311. zlRecalItemObjectRowNo:根据结算列表信息,重新刷新结算对象的行号属性
'    0312. zlSetBalanceRowDataFromItemsObject:根据结算信息，设置相关结算列表中的行数据
'    0313. zlSetBalanceRowDataFromItemObject:根据结算项，设置指定行的结算状态
'    0314. zlGetBalanceCancelSQL:获取结帐取消保存的相关SQL
'    0315. zlMoveRowBalanceFromSwapID:根据关联交易ID,删除对应的结算列表中对应的行
'    0317. zlReCalcBalanceInfor:重新计算结帐信息的未付及已付
'    0318. zlCheckMulitInterfaceNumValied:检查是正同时存在三种以上接口(不含三种)
'    0319. zlGetPtBalanceItemsFromVsBalance:获取普通的结算信息对象集
'    0320. zlSetVsBalanceEditStatus:设置结算的编译状态
'    0321. zlGetLedDisplayBankDatasFromVsBalance:根据结算列表，获取显示在Led上的结算数据集
'    0322. zlGetBalanceIDFromBalanceNo-获取原结帐ID
'  04:预交相关
'    0400. zlInitDepositGrid:初始化预交网格
'    0401. zlGetDelDepositItemsFromVsDeposit:根据预交列表，获取退款信息集
'    0402. zlGetThirdTransferItemsFromVsDeposit:根据预交款列表数据（转帐）,获取三方退款的分摊明细信息
'    0403. zlRecalcDepositMoney：重新计算冲预交金额
'    0405. zlLoadDepositListFromBalanceID:根据结帐ID获取冲预交信息信息并加载到预交列表中
'    0406. zlLoadDepositListFromRecord:根据预交记录集，将预交信息加载到预交列表中
'    0407. zlGetThridTransItemsFromVsDepositAndTranItem:根据转账金额，重新获取需要转帐的子项
'    0408. zlGetItemFromVsDepositRow:根据预交款行，获取该项的结算信息
'  05.费用相关
'    0501:zlAutoRecalFeeBalanceMoney：自动计算或分摊结帐金额
'    0502.zlLoadDetaiFeeToGridFromRecord:根据费用记录集，将数据加载到费用明细网格中
'    0503:zlLoadDetaiFeeToGridFromBalanceID:根据结帐ID,将费用加载到费用明细网格中
'    0504:zlLoadFeiMuFeeListToGridFromRecord:根据费用记录集，将数据加载到费目明细网格中
'    0505:zlLoadFeiMuFeeListToGridFromBalanceID:根据结帐ID,将费目费用加载到费目明细网格中
'    0506. zlGetReadFeeDetailFromBalanceID:根据结帐DI,获取主界面费用明细列表数据
'    0507. zlGetExceptionBalanceData:获取异常的结帐数据根据时间范围和操作员姓名
'  06:Items对象操作相关
'    0601. zlCopyNewItemFromBalanceItem:复制一个新的Item对象
'编制:刘兴洪
'日期:2018-05-23 14:40:18
'*********************************************************************************************************************************************
Public grs结算方式 As ADODB.Recordset
Public Const g_BalanceRow_Color_Succes = &H80000011  '接口调用成功:灰色
Public Const g_BalanceRow_Color_Valied = &HFF&       '接口调用失败:红色
Public Const g_BalanceRow_Color_Normal = &H80000008  '正常的查看:黑色



Public Function zlCopyNewFeeData(ByVal bytOperation As Byte, ByVal rsFeeList As ADODB.Recordset, ByRef rsNewFeeList_Out As ADODB.Recordset, Optional strOwnerFeeType As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一个记录集，拷贝为一个新的记录集
    '入参:bytOperation-0-拷贝自费费用数据;1--拷贝血费费用数据
    '     rsFeeList-原记录数据
    '     strOwnerFeeType-自费类型,多个用逗号分离
    '出参:rsNewFeeList_Out-返回新的符合条件数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-10-29 16:08:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, varData As Variant, i As Long
    On Error GoTo errHandle
    
    If bytOperation = 0 Then '自费部分
        If varData = "" Then Set rsNewFeeList_Out = rsFeeList: zlCopyNewFeeData = True: Exit Function
        varData = Split(strOwnerFeeType, ",")
        For i = 0 To UBound(varData)
            strFilter = strFilter & " Or 收费类别='" & Replace(varData(i), "'", "") & "'"
        Next
        strFilter = Mid(strFilter, 4)
        rsFeeList.Filter = strFilter
        Set rsNewFeeList_Out = zlDatabase.CopyNewRec(rsFeeList)
        rsFeeList.Filter = 0
        zlCopyNewFeeData = True: Exit Function
    End If
    
    If bytOperation = 1 Then '血库部分
        strFilter = " 收费类别='K'"
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
    '功能:比较两个住院次数是否一致
    '入参:strPatiNums1-病人住院次数1
    '     strPatiNums2-病人住院次数2
    '出参:
    '返回:一致返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-10-29 17:12:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, i As Long, j As Long, blnFind As Boolean

    On Error GoTo errHandle
    If strPatiNums1 = strPatiNums2 Then zlComparePatiNumsIsDiff = True: Exit Function '相同的值，肯定一致
    
    varData = Split(strPatiNums1, ","): varTemp = Split(strPatiNums2, ",")
    If UBound(varData) <> UBound(varTemp) Then zlComparePatiNumsIsDiff = False: Exit Function  '住院次数不一样，肯定一致(即始包含了一样的，也判断为不一致)
    
    For i = 0 To UBound(varData)
        blnFind = False
        For j = 0 To UBound(varTemp)
            If Val(varData(i)) = Val(varTemp(j)) Then
                blnFind = True: Exit For
            End If
        Next
        If Not blnFind Then zlComparePatiNumsIsDiff = False: Exit Function  '未找到，肯定不一致
    Next
    For i = 0 To UBound(varTemp)
        blnFind = False
        For j = 0 To UBound(varData)
            If Val(varTemp(i)) = Val(varData(j)) Then
                blnFind = True: Exit For
            End If
        Next
        If Not blnFind Then zlComparePatiNumsIsDiff = False: Exit Function  '未找到，肯定不一致
    Next
    zlComparePatiNumsIsDiff = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function zlGetBalanceDataErrFromBalanceID(ByVal lng结帐ID As Long, ByRef rsBalance_Out As ADODB.Recordset, _
    Optional blnDel As Boolean, Optional blnMoved As Boolean, Optional strTittle As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID获取相关的异常结帐数据信息
    '入参:lng结帐ID-结帐ID
    '     strCaptions-标题名称
    '     blnMoved-是否进行了历史数据转移
    '出参:rsBalance_Out-获取成功时，返回的结帐信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-05-23 14:42:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    strTittle = IIf(strTittle = "", "根据结帐ID获取相关的异常结帐数据信息（zlGetBalanceDataErrFromBalanceID)", strTittle)
    strSQL = " " & _
    "    Select 结算方式, Sum(冲预交) As 冲预交, 标志, 性质 " & _
    "    From (Select Decode(Mod(记录性质, 10), 1, '[冲预交]', Nvl(a.结算方式, '未结金额')) As 结算方式, " & IIf(blnDel, "-1*", "") & " a.冲预交 as 冲预交, " & _
    "                Decode(Nvl(a.校对标志, 0), 0, '√', 2, '√', '×') As 标志, Decode(Mod(记录性质, 10), 1, -1, Nvl(b.性质, 0)) As 性质 " & _
    "           From 病人预交记录 A, 结算方式 B " & _
    "           Where a.结帐id = [1] And a.结算方式 = b.名称(+)) A " & _
    "    Group By 结算方式, 标志, 性质 " & _
    "    Having Sum(a.冲预交) <> 0 " & _
    "    Order By 性质"
    
    If blnMoved Then
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
    End If
    Set rsBalance_Out = zlDatabase.OpenSQLRecord(strSQL, strTittle, lng结帐ID)
    zlGetBalanceDataErrFromBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceDataFromBalanceID(ByVal lng结帐ID As Long, ByRef rsBalance_Out As ADODB.Recordset, _
    Optional blnDel As Boolean, Optional blnMoved As Boolean, Optional strTittle As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID获取相关的结帐数据信息
    '入参:lng结帐ID-结帐ID
    '     strCaptions-标题名称
    '     blnMoved-是否进行了历史数据转移
    '出参:rsBalance_Out-获取成功时，返回的结帐信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-05-23 14:42:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    strTittle = IIf(strTittle = "", "根据结帐ID获取相关的结帐数据信息（zlGetBalanceDataFromBalanceID)", strTittle)
    
    strSQL = _
    " Select Decode(Mod(记录性质, 10), 1,'冲预交',decode(结算方式,NULL,'未结','补款')) as 类型,NO as 单据号," & IIf(blnDel, "-1*", "") & "冲预交 as 金额," & _
    "       结算方式,结算号码,是否电子票据  " & _
    " From 病人预交记录  " & _
    " Where 结帐ID=[1] And 冲预交 <> 0 " & _
    " Order by 类型 Desc,NO Desc,结算方式"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
    End If
    
    Set rsBalance_Out = zlDatabase.OpenSQLRecord(strSQL, strTittle, lng结帐ID)
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
    '功能:根据原交易信息获取保存三方退款信息的SQL
    '入参:objItem-需要批量退款的三方结算信息
    '     blnRetrunXML-是否返回XML串
    '出参:cllPro-返回的SQL集
    '     strInXml_Out-blnRetrunXML=true时，返回,格式为:
    '     objItems_Out-返回结算退款信息明细
    '     strBalanceDepsoitIDs-预交ID集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-05-21 11:02:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    Dim objItemTemp As clsBalanceItem
    Dim i As Long
    
    On Error GoTo errHandle
    
    strInXml_Out = "": strDepsoitIDs = ""
    ' 卡号,交易流水号,交易说明,金额,预交ID
    varData = Split(objItem.Tag, "|")
    Set objItems_Out = New clsBalanceItems
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",,,,", ",")
        Set objItemTemp = New clsBalanceItem
        With objItemTemp
            Set .objCard = objItem.objCard
            .卡号 = varTemp(0)
            .交易流水号 = varTemp(1)
            .交易说明 = varTemp(2)
            .结算金额 = RoundEx(-1 * Val(varTemp(3)), 2)
            .结帐ID = objItem.结帐ID
            .行号 = objItem.行号
            .结帐时间 = objItem.结帐时间
            .结算方式 = objItem.结算方式
            .卡类别ID = objItem.卡类别ID
            .门诊结帐 = objItem.门诊结帐
            .是否退款分交易 = objItem.是否退款分交易
            .是否转帐 = objItem.是否转帐
            .校对标志 = objItem.校对标志
            .预交ID = Val(varTemp(4))
        End With
        
        objItems_Out.AddItem objItemTemp
        If blnRetrunXML Then
            strInXml_Out = strInXml_Out & "<JS>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <KH>" & objItemTemp.卡号 & "</KH>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <JYLSH>" & TruncStringEx(objItemTemp.交易流水号, True) & "</JYLSH>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <JYSM>" & TruncStringEx(objItemTemp.交易流水号, True) & "</JYSM>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <ZFJE>" & objItemTemp.结算金额 & "</ZFJE>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <JSLX>" & 1 & "</JSLX>" & vbCrLf
            strInXml_Out = strInXml_Out & "     <ID>" & objItemTemp.预交ID & "</ID>" & vbCrLf
            strInXml_Out = strInXml_Out & "</JS>" & vbCrLf
        End If
        
        If zlGetSaveThirdSwapDelSQLFromBalanceItem(objItemTemp, True, cllPro) = False Then Exit Function
        strDepsoitIDs = strDepsoitIDs & "," & objItemTemp.预交ID
    Next i
    If strDepsoitIDs <> "" Then strDepsoitIDs = Mid(strDepsoitIDs, 2)
    zlGetSaveThreeDelSwapBatchSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetSaveThirdSwapDelSQLFromBalanceItem(ByVal objItem As clsBalanceItem, ByVal blnModify As Boolean, ByRef cllPro As Collection, Optional int校对标志 As Integer = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算对象获取三方交易退款保存SQL
    '入参:objItem-当前结帐对象
    '     blnModify-是否修改
    '     bln转帐-是否当前进行的转帐操作
    '     int校对标志-1-接口未成功;0-接口调用成功
    '出参:cllPro-返回的结算信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-05-20 18:08:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    ' Zl_三方退款信息_Insert
    strSQL = "Zl_三方退款信息_Insert("
    '  结帐id_In     三方退款信息.结帐id%Type,
    strSQL = strSQL & "" & objItem.结帐ID & ","
    '  记录id_In     三方退款信息.记录id%Type,
    strSQL = strSQL & "" & objItem.预交ID & ","
    '  金额_In       三方退款信息.金额%Type,
    strSQL = strSQL & "" & Abs(objItem.结算金额) & ","
    '  卡号_In       三方退款信息.卡号%Type,
    strSQL = strSQL & "'" & objItem.卡号 & "',"
    '  交易流水号_In 三方退款信息.交易流水号%Type,
    strSQL = strSQL & "'" & objItem.退款交易流水号 & "',"
    '  交易说明_In   三方退款信息.交易说明%Type,
    strSQL = strSQL & "'" & objItem.退款交易说明 & "',"
    '  操作类型_In   Number := 0,
    strSQL = strSQL & "'" & IIf(blnModify, 1, 0) & "',"
    '  是否未退_In   三方退款信息.是否未退%Type := 0
    strSQL = strSQL & "'" & IIf(int校对标志 = 1, 1, 0) & "',"
    '  是否转帐_In   三方退款信息.是否转帐%Type := 0
    strSQL = strSQL & "'" & IIf(objItem.是否转帐, 1, 0) & "',"
    '  卡类别id_In   三方退款信息.卡类别id%Type := Null
    strSQL = strSQL & "" & IIf(objItem.卡类别ID = 0, "NULL", objItem.卡类别ID) & ","
    '  原交易流水号_In 三方退款信息.原交易流水号%Type := Null,
    strSQL = strSQL & "'" & objItem.交易流水号 & "',"
    '  原交易说明_In   三方退款信息.原交易说明%Type := Null
    strSQL = strSQL & "'" & objItem.交易说明 & "')"
  
    zlAddArray cllPro, strSQL
    zlGetSaveThirdSwapDelSQLFromBalanceItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetSaveThirdSwapDelSQLFromBalanceItems(ByVal objItems As clsBalanceItems, ByVal blnModify As Boolean, _
    ByRef cllPro As Collection, Optional int校对标志 As Integer = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算对象获取三方交易退款保存SQL
    '入参:objItem-当前结帐对象
    '     blnModify-是否修改
    '出参:cllPro-返回的结算信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-05-20 18:08:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim objItem As clsBalanceItem
    
    On Error GoTo errHandle
    For Each objItem In objItems
        If zlGetSaveThirdSwapDelSQLFromBalanceItem(objItem, blnModify, cllPro, int校对标志) = False Then Exit Function
    Next
    zlGetSaveThirdSwapDelSQLFromBalanceItems = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlThirdDelSwapIsExsistFromBalanceID(ByVal lng结帐ID As Long, Optional bln含未退 As Boolean = True, Optional strTittle As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID，判断三方退款交易是否存在
    '入参:bln含未退-是否包含未退部分的检查
    '出参:
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-05-25 09:55:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strTittle = IIf(strTittle = "", "根据结帐ID，判断三方退款交易是否存在（zlThirdDelSwapIsExsistFromBalanceID)", strTittle)
 
    strWhere = ""
    If Not bln含未退 Then strWhere = " And nvl(是否未退,0)<>1"
    strSQL = "Select 1 From 三方退款信息 Where 结帐ID=[1] And Rownum<2 " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, strTittle, lng结帐ID)
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
    '功能:根据预交列表来获取转帐项
    '入参:objTranItem-转帐项
    '     dblTransMoney-转帐金额:0表示获取相同的所有项信息
    '     strNotCardTypeIDs-不包含的卡类别,多个用逗号分离
    '出参:objItems_Out-返回相关子项
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-27 11:59:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, bln消费卡 As Boolean, lngCardTypeID As Long, lngCurCardTypeID As Long
    Dim objItem As clsBalanceItem, objCard As Card
    Dim dblTranTotal As Double, dbl冲预交 As Double, dblMoney As Double
    Dim blnAll As Boolean, i As Long, bln转帐 As Boolean
    Dim strCardOwerCardtypeIDs As String
    
    On Error GoTo errHandle
    lngCurCardTypeID = 0
    If Not objTranItem Is Nothing Then
        Set objCard = objTranItem.objCard
        lngCurCardTypeID = objTranItem.卡类别ID
    End If
    
    blnAll = dblTransMoney = 0
    dblTranTotal = dblTransMoney
    If objItems_Out Is Nothing Then Set objItems_Out = New clsBalanceItems
    For i = 1 To objItems_Out.Count
        If InStr("," & strCardOwerCardtypeIDs & ",", "," & objItems_Out(i).卡类别ID & ",") = 0 Then
            strCardOwerCardtypeIDs = strCardOwerCardtypeIDs & "," & objItems_Out(i).卡类别ID
        End If
    Next
    
    With vsDeposit
        For i = .Rows - 1 To 1 Step -1
            
            lngCardTypeID = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
 
            strNO = Trim(.TextMatrix(i, .ColIndex("单据号")))
            bln消费卡 = Val(.TextMatrix(i, .ColIndex("是否消费卡"))) = 1
            dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
 
            bln转帐 = Val(.TextMatrix(i, .ColIndex("是否转帐及代扣"))) = 1
            If dblTransMoney = 0 And Not blnAll Then zlGetThridTransItemsFromVsDepositAndTranItem = True: Exit Function
            
            '要支持转帐的三方卡，才能合并转帐
            If ((lngCurCardTypeID = lngCardTypeID And bln消费卡 = False) Or lngCurCardTypeID = 0 Or (bln转帐 And lngCardTypeID <> 0 And lngCurCardTypeID = 0 And bln消费卡 = False)) _
                And InStr("," & strNotCardTypeIDs & strCardOwerCardtypeIDs & ",", "," & lngCardTypeID & ",") = 0 Then
                
                If dblTransMoney > dbl冲预交 Or blnAll Then
                    dblMoney = dbl冲预交
                    If Not blnAll Then dblTransMoney = RoundEx(dblTransMoney - dbl冲预交, 6)
                Else
                    dblMoney = dblTransMoney
                    dblTransMoney = 0
                End If
                
                If lngCurCardTypeID <> 0 Then
                    Set objItem = zlCopyNewItemFromBalanceItem(objTranItem)
                    If objCard Is Nothing Then Set objCard = zlGetCardFromCardType(lngCardTypeID, False, Trim(.TextMatrix(i, .ColIndex("结算方式"))))
                Else
                    Set objItem = New clsBalanceItem
                    Set objCard = zlGetCardFromCardType(lngCardTypeID, False, Trim(.TextMatrix(i, .ColIndex("结算方式"))))
                End If
                Set objItem.objCard = objCard
                objItem.结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
                objItem.关联交易ID = Val(.TextMatrix(i, .ColIndex("关联交易ID")))
                objItem.卡类别ID = lngCardTypeID
                objItem.卡号 = Trim(.TextMatrix(i, .ColIndex("卡号")))
                objItem.交易流水号 = Trim(.TextMatrix(i, .ColIndex("交易流水号")))
                objItem.交易说明 = Trim(.TextMatrix(i, .ColIndex("交易说明")))
                objItem.结算号码 = Trim(.TextMatrix(i, .ColIndex("结算号码")))
                objItem.结算金额 = RoundEx(-1 * dblMoney, 6)
                objItem.结算摘要 = Trim(.TextMatrix(i, .ColIndex("摘要")))
                objItem.门诊结帐 = IIf(objBalanceInfor.结算类型 = 1, True, False)
                objItem.冲销ID = objBalanceInfor.冲销ID
                objItem.结算IDs = objBalanceInfor.结帐ID
                objItem.结帐ID = objBalanceInfor.结帐ID
                objItem.结帐时间 = objBalanceInfor.结帐时间
                objItem.结算性质 = objCard.结算性质
                objItem.是否预交 = True
                objItem.是否退款 = True
                If lngCardTypeID <> 0 Then
                    objItem.结算类型 = IIf(bln消费卡, 5, 3) '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                ElseIf objCard.结算性质 = 7 Then
                    objItem.结算类型 = 4
                Else
                    objItem.结算类型 = 0
                End If
                objItem.预交ID = Val(.TextMatrix(i, .ColIndex("预交ID")))
                objItem.是否密文 = objCard.卡号密文规则 <> ""
                objItem.是否转帐 = True
                objItem.校对标志 = 1
             
                objItems_Out.AddItem objItem
                objItems_Out.结算金额 = RoundEx(objItems_Out.结算金额 + objItem.结算金额, 6)
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
    '功能:获取预交款的三方退款的分摊明细信息
    '入参:vsDeposit-预交网格数据
    '     objDelItems-当前退款信息集
    '     objTranItem-当前的转帐项目
    '出参:objTranItem_Out-当前返回的转帐信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-06-14 15:14:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, objTempItem As clsBalanceItem
    Dim objItems As clsBalanceItems, dblMoney As Double
    Dim strDelCardTypeIDs As String, dblBalance As Double

    If objCurTranItem Is Nothing Then zlGetThirdTransferItemsFromVsDeposit = True: Exit Function
    
    '正式处理转帐交易数据
    For Each objItem In objDelItems
        strDelCardTypeIDs = strDelCardTypeIDs & "," & objItem.卡类别ID
    Next
    
    Err = 0: On Error GoTo errHandle
    
    If objCurTranItem Is Nothing Then Exit Function
    
    dblMoney = -1 * objCurTranItem.结算金额
    '第一步:先处理自身的转帐
    If zlGetThridTransItemsFromVsDepositAndTranItem(vsDeposit, objBalanceInfor, strDelCardTypeIDs, objCurTranItem, objItems, dblMoney) = False Then Exit Function
    
    If RoundEx(dblMoney, 6) = 0 Then
        If objItems Is Nothing Then Exit Function
        If objItems.Count = 0 Then Exit Function
        Set objCurTranItem.objTag = objItems
        Set objTranItem_Out = objCurTranItem
        objTranItem_Out.结算金额 = objCurTranItem.结算金额
        zlGetThirdTransferItemsFromVsDeposit = True: Exit Function
    End If
    
    '第二步：再处理其他的转帐(剩余金额分摊)
    If zlGetThridTransItemsFromVsDepositAndTranItem(vsDeposit, objBalanceInfor, strDelCardTypeIDs, Nothing, objItems, dblMoney) = False Then Exit Function
    If objItems Is Nothing Then Exit Function
    If objItems.Count = 0 Then Exit Function
    If RoundEx(dblMoney, 6) <> 0 Then
        dblBalance = RoundEx(objBalanceInfor.当前结帐 - objBalanceInfor.医保支付合计 - objBalanceInfor.误差费, 6)
        If RoundEx(dblBalance, 6) < 0 Then dblBalance = 0 '医保报销可能大于费用总金额
        MsgBox "你当前转帐金额大于了预存款应退的金额,请检查!" & vbCrLf & _
               "转帐金额:" & Format(-1 * objCurTranItem.结算金额, "0.00") & vbCrLf & _
               "应退金额:" & Format(objBalanceInfor.冲预交合计 - dblBalance, "0.00"), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    For Each objItem In objItems
        objItem.卡类别ID = objCurTranItem.卡类别ID
    Next
    Set objCurTranItem.objTag = objItems
    Set objTranItem_Out = objCurTranItem
    objTranItem_Out.结算金额 = objCurTranItem.结算金额
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
    '功能:获取预交款的三方退款的分摊明细信息
    '入参:objThirdSwap-三方交易接口对象
    '       vsDeposit-预交网格数据
    '       dblNotFeeTotal-未结费用总额(此项金额费用总额-医保支付总额)
    '出参:objItems_Out-获取退款信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-06-14 15:14:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblBalanceSum As Double, dblTemp As Double, dbl余额 As Double, dbl原始金额 As Double, dbl冲预交 As Double, dblMoney As Double, dblDelMoney As Double
    Dim strSwapNO As String, strSwapDemo As String, strNO As String, str卡号 As String, strSQL As String, str结算方式 As String, strDefaultBalance As String
    Dim strErrMsg As String, strExpend As String, bln消费卡 As Boolean, blnAdd As Boolean, bln是否转帐 As Boolean, blnDelCash As Boolean, blnFind As Boolean
    Dim lng预交ID As Long, lngCardTypeID As Long, lng序号 As Long, i As Long, j As Long, lng关联交易ID As Long, int结算性质 As Integer
    Dim objItemsPt As clsBalanceItems, objOldItems As clsBalanceItems
    Dim objItems As clsBalanceItems, objItemsTemp As clsBalanceItems, objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim objData As clsBalanceData, objDatasTemp As clsBalanceDatas
    Dim objDataMulit As clsBalanceDatas '多个交易集
    Dim objDataSingle As clsBalanceDatas  '单一集
    Dim objDataTrans As clsBalanceDatas '转帐集
    Dim blnSingleDel As Boolean '是否单交易
    Dim cllDelSwap As Collection 'array(卡类别ID,是否调用一次接口交易) ：是否调用一次接口交易:1-是;0-否)
    Dim rsTemp As ADODB.Recordset
    Dim varData As Variant, blnDelToLocalMode As Boolean  '结帐时,有预交剩余款时,预交剩余款是否退到指定结算方式
    Dim str结算号码 As String, str结算摘要 As String
    
    dblBalanceSum = RoundEx(dblDelTotal, 2) '结帐金额可能有多位，因此，只能四舍五入到2位来进行处理
    If dblBalanceSum <= 0 Then zlGetDelDepositItemsFromVsDeposit = True: Exit Function
    blnDelToLocalMode = gTy_System_Para.TY_Balance.bln预交退指定结算方式 And gTy_System_Para.TY_Balance.str预交退款结算方式 <> ""
    '初始化数据结构
    Set rsTemp = New ADODB.Recordset
    rsTemp.Fields.Append "序号", adInteger, , adFldIsNullable
    rsTemp.Fields.Append "性质", adDouble, , adFldIsNullable    '0-三方卡;1-现金,2-消费卡等
    rsTemp.Fields.Append "卡类别ID", adVarChar, 50, adFldIsNullable
    rsTemp.Fields.Append "是否消费卡", adInteger, , adFldIsNullable
    rsTemp.Fields.Append "关联交易ID", adVarChar, 50, adFldIsNullable
    rsTemp.Fields.Append "交易流水号", adVarChar, 100, adFldIsNullable
    rsTemp.Fields.Append "结算方式", adVarChar, 50, adFldIsNullable
    rsTemp.Fields.Append "原始金额", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "冲预交", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "剩余金额", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "余额", adDouble, , adFldIsNullable
 
    rsTemp.CursorLocation = adUseClient
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorType = adOpenStatic
    rsTemp.Open
    
    
    On Error GoTo errHandle
    lng序号 = 0
    Set cllDelSwap = New Collection
    '1.先汇总各项结算方式的总金额(原因是有正有负。可能同一笔交易流水号已经余额退未或退款，需要排除)
    With vsDeposit
           For i = 1 To .Rows - 1
                lngCardTypeID = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                strSwapNO = Trim(.TextMatrix(i, .ColIndex("交易流水号")))
                bln消费卡 = Val(.TextMatrix(i, .ColIndex("是否消费卡"))) = 1
                dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
                lng预交ID = Val(.TextMatrix(i, .ColIndex("预交ID")))
                lng关联交易ID = Val(.TextMatrix(i, .ColIndex("关联交易ID")))
                bln是否转帐 = Val(.TextMatrix(i, .ColIndex("是否转帐及代扣")))
                str卡号 = Val(.TextMatrix(i, .ColIndex("卡号")))
                dbl余额 = Val(.TextMatrix(i, .ColIndex("余额")))
                dbl原始金额 = Val(.TextMatrix(i, .ColIndex("原始金额")))
                strNO = .TextMatrix(i, .ColIndex("单据号"))
                
                If strNO <> "" Then
                    If blnDelToLocalMode Then   '预交退款退到指定结算方式，即始是三方卡，也不调接口
                        dblMoney = RoundEx(dblMoney + dbl冲预交, 6)
                    Else
                        If strSwapNO = "" Then strSwapNO = " "
                        If lngCardTypeID <> 0 Then
                            If bln消费卡 Then
                                rsTemp.Filter = "卡类别ID=" & lngCardTypeID
                            Else
                                If lng关联交易ID <> 0 Then
                                     rsTemp.Filter = "卡类别ID=" & lngCardTypeID & " and 关联交易ID=" & lng关联交易ID
                                Else
                                     rsTemp.Filter = "卡类别ID=" & lngCardTypeID & " and 交易流水号='" & strSwapNO & "'"
                                End If
                            End If
                            If rsTemp.EOF Then rsTemp.AddNew: lng序号 = lng序号 + 1: rsTemp!序号 = lng序号
                        Else
                            '非三方卡的，直接处理
                            lng序号 = lng序号 + 1
                            If lng关联交易ID <> 0 Then
                                rsTemp.Filter = "卡类别ID=" & lngCardTypeID & " and 关联交易ID=" & lng关联交易ID
                            Else
                                rsTemp.Filter = "卡类别ID=" & lngCardTypeID & " and 是否消费卡=" & IIf(bln消费卡, 1, 0) & " and 结算方式='" & Trim(.TextMatrix(i, .ColIndex("结算方式"))) & "'"
                            End If
                            If rsTemp.EOF Then rsTemp.AddNew: lng序号 = lng序号 + 1: rsTemp!序号 = lng序号
                        End If
                        rsTemp!性质 = IIf(lngCardTypeID <> 0, IIf(bln消费卡, 2, 0), 1)
                        rsTemp!卡类别ID = lngCardTypeID
                        rsTemp!关联交易ID = lng关联交易ID
                        rsTemp!结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
                        rsTemp!剩余金额 = NVL(rsTemp!剩余金额, 0) + dbl冲预交
                        rsTemp!结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
                        rsTemp!余额 = Val(NVL(rsTemp!余额, 0)) + dbl余额
                        rsTemp!是否消费卡 = IIf(bln消费卡, 1, 0)
                        rsTemp!原始金额 = Val(NVL(rsTemp!原始金额)) + IIf(dbl原始金额 > 0, dbl原始金额, Val(NVL(rsTemp!原始金额)))
                        
                        rsTemp.Update
                    End If
                End If
           Next
    End With
    
    If blnDelToLocalMode Then   '预交退款退到指定结算方式，即始是三方卡，也不调接口
        dblMoney = dblMoney - dblNotFeeTotal
        Call zlAddfinancialTrancsToBalanceList(vsBlance, -1 * dblMoney)
        zlGetDelDepositItemsFromVsDeposit = True
        Exit Function
    End If
    
    Set objDataMulit = New clsBalanceDatas '多个交易集
    Set objDataSingle = New clsBalanceDatas '单一集
    Set objDataTrans = New clsBalanceDatas '转帐集
    
    Set objItemsPt = New clsBalanceItems
    Set objItems = New clsBalanceItems
    Set objOldItems = New clsBalanceItems
    dblMoney = 0
    '计算三方退款信息或转帐信息
     With vsDeposit
        For i = .Rows - 1 To 1 Step -1
            If dblBalanceSum > 0 Then
                
                lng预交ID = Val(.TextMatrix(i, .ColIndex("预交ID")))
                dblMoney = Val(.TextMatrix(i, .ColIndex("冲预交")))
                lngCardTypeID = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                lng关联交易ID = Val(.TextMatrix(i, .ColIndex("关联交易ID")))
                bln消费卡 = Val(.TextMatrix(i, .ColIndex("是否消费卡"))) = 1
                int结算性质 = Val(.TextMatrix(i, .ColIndex("结算性质")))
                str结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
                bln是否转帐 = Val(.TextMatrix(i, .ColIndex("是否转帐及代扣"))) = 1
                strSwapNO = Trim(.TextMatrix(i, .ColIndex("交易流水号")))
                
                If dblMoney < 0 Then GoTo GoNext   '退款时，直接忽略
                
                If lngCardTypeID <> 0 Then
                   If bln消费卡 Then
                       rsTemp.Filter = "卡类别id=" & lngCardTypeID & " And 是否消费卡=1 And 剩余金额>0 "
                       If rsTemp.EOF Then GoTo GoNext
                       
                       dblTemp = Val(NVL(rsTemp!剩余金额))
                       If dblTemp = 0 Then GoTo GoNext
                       
                       If dblTemp < dblMoney Then dblMoney = dblTemp '本次只能冲剩余金额
                       
                       If dblBalanceSum > dblMoney Then
                           dblBalanceSum = RoundEx(dblBalanceSum - dblMoney, 6)
                           dblDelMoney = dblMoney
                       Else
                           'If objItem.objCard.是否全退 Then dblDelMoney = Val(NVL(rsTemp!剩余金额))    '必须全退
                           dblDelMoney = dblBalanceSum: dblBalanceSum = 0
                       End If
                       If dblDelMoney = 0 Then GoTo GoNext
                       
                
                       dbl原始金额 = Val(NVL(rsTemp!原始金额))
                       
                       Set objItem = New clsBalanceItem
                       With objItem
                           Set .objCard = zlGetCardFromCardType(lngCardTypeID, bln消费卡, str结算方式)
                           .是否转帐 = False
                           .结算性质 = int结算性质
                           .结算IDs = ""
                           .交易流水号 = strSwapNO
                           .交易说明 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("交易说明")))
                           .卡号 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("卡号")))
                           .关联交易ID = 0
                           .结算金额 = -1 * dblDelMoney
                           .结算方式 = str结算方式
                           .结算号码 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("结算号码")))
                           .结算摘要 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("摘要")))
                           .结算类型 = 5    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                           .是否预交 = True
                           .是否退款 = True
                           .是否允许退现 = .objCard.是否退现
                           .是否允许编辑 = False
                           .是否允许删除 = .是否允许退现
                           .未退金额 = Val(NVL(rsTemp!余额))
                           .原始金额 = dbl原始金额
                           .卡类别ID = lngCardTypeID
                           .消费卡 = bln消费卡
                           .是否密文 = IIf(.objCard.卡号密文规则 <> "", True, False)
                           .结帐时间 = CDate(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("收款日期")))
                           .预交ID = lng预交ID
                       End With
                       
                       Set objData = New clsBalanceData
                       Set objItemsTemp = New clsBalanceItems
                       objItemsTemp.AddItem objItem
                       objItemsTemp.结算金额 = RoundEx(objItemsTemp.结算金额 + objItem.结算金额, 6)
                       objItemsTemp.收费类型 = 1
                       objItemsTemp.是否转帐 = False
                       objData.Key = "K" & lngCardTypeID & "_1"
                       Set objData.objBalanceItems = objItemsTemp
                       objDataSingle.AddItem objData
                   Else
                       '三方卡
                       If lng关联交易ID <> 0 Then
                           rsTemp.Filter = "卡类别id=" & lngCardTypeID & " And 关联交易ID=" & lng关联交易ID & " And 剩余金额>0 "
                       Else
                            If strSwapNO = "" Then strSwapNO = " "
                            rsTemp.Filter = "卡类别ID=" & lngCardTypeID & " and 交易流水号='" & strSwapNO & "' And 剩余金额>0"
                       End If
                       
                       If rsTemp.EOF Then GoTo GoNext
                       dblTemp = Val(NVL(rsTemp!剩余金额))
                       If dblTemp = 0 Then GoTo GoNext
                       
                       If dblTemp < dblMoney Then dblMoney = dblTemp '本次只能冲剩余金额
                       
                       If dblBalanceSum > dblMoney Then
                           dblBalanceSum = RoundEx(dblBalanceSum - dblMoney, 6)
                           dblDelMoney = dblMoney
                       Else
                           'If objItem.objCard.是否全退 Then dblDelMoney = Val(NVL(rsTemp!剩余金额))    '必须全退
                           dblDelMoney = dblBalanceSum: dblBalanceSum = 0
                       End If
                       
                       If dblDelMoney = 0 Then GoTo GoNext
                       
                
                       dbl原始金额 = Val(NVL(rsTemp!原始金额))
                       
                       Set objItem = New clsBalanceItem
                       With objItem
                           Set .objCard = zlGetCardFromCardType(lngCardTypeID, bln消费卡, str结算方式)
                           .objCard.是否转帐及代扣 = bln是否转帐
                           .是否转帐 = bln是否转帐
                           .结算性质 = int结算性质
                           .结算IDs = ""
                           .交易流水号 = strSwapNO
                           .交易说明 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("交易说明")))
                           .卡号 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("卡号")))
                           .关联交易ID = lng关联交易ID
                           .结算金额 = -1 * dblDelMoney
                           .结算方式 = str结算方式
                           .结算号码 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("结算号码")))
                           .结算摘要 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("摘要")))
                           .结算类型 = IIf(.结算性质 = 7, 4, IIf(Not bln消费卡, 3, 5)) '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                           .是否预交 = True
                           .是否退款 = True
                           .是否允许编辑 = False
                           .是否允许删除 = True
                           .未退金额 = Val(NVL(rsTemp!余额))
                           .原始金额 = dbl原始金额
                           .卡类别ID = lngCardTypeID
                           .消费卡 = bln消费卡
                           .是否密文 = IIf(.objCard.卡号密文规则 <> "", True, False)
                           .结帐时间 = CDate(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("收款日期")))
                           .是否允许退现 = .objCard.是否退现
                           .预交ID = lng预交ID
                           .消费卡ID = 0
                           .是否退款分交易 = True
                       End With
                       
                       Set objItemsTemp = New clsBalanceItems
                       
                       objItemsTemp.AddItem objItem
                       objItemsTemp.结算金额 = objItem.结算金额
                       objItemsTemp.收费类型 = 1
                       
                       blnAdd = False
                       If objItem.objCard.是否转帐及代扣 Then
                           '转帐及代扣，不需要调用接口
                           blnAdd = True
                       Else
                           If Not objThirdSwap.zlThirdReturnCashCheck(objItem.objCard, objItemsTemp, blnDelCash, strDefaultBalance) Then
                               '1.禁止退现
                               objItem.是否允许退现 = False
                               objItem.是否强制退现 = blnDelCash
                               objItem.是否允许删除 = objItem.是否强制退现
                               blnAdd = True
                           Else
                               If blnDelCash = False Then  '是否缺省退现
                                   '允许退现，可以删除
                                   objItem.是否允许编辑 = False
                                   objItem.是否允许删除 = True
                                   objItem.是否强制退现 = True
                                   objItem.是否允许退现 = True: blnAdd = True
                               ElseIf strDefaultBalance <> "" Then
                               
                                   blnFind = False
                                   For j = 1 To objItemsPt.Count
                                       If objItemsPt(j).结算方式 = strDefaultBalance Then
                                           objItemsPt(j).结算金额 = objItemsPt(j).结算金额 + objItem.结算金额
                                           objItemsPt.结算金额 = objItemsPt.结算金额 + objItem.结算金额
                                           blnFind = True
                                           Exit For
                                       End If
                                   Next
                                   
                                   If Not blnFind Then
                                       Set objItem = New clsBalanceItem
                                       With objItem
                                           Set .objCard = zlGetCardFromBalanceName(strDefaultBalance)
                                           .结算方式 = strDefaultBalance
                                           .结算金额 = RoundEx(-1 * dblDelMoney, 6)
                                           .是否退款 = True
                                           .是否允许编辑 = False
                                           .是否允许删除 = True
                                           .结算性质 = .objCard.结算性质
                                           .结算类型 = 0 '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                                           
                                           .Tag = "指定预交退款"
                                       End With
                                       objItemsPt.AddItem objItem
                                       objItemsPt.结算金额 = RoundEx(objItemsPt.结算金额 + objItem.结算金额, 6)
                                   End If
                                   
                               End If
                           End If
                       End If
                       
                       If blnAdd Then
                           If objItem.objCard.是否转帐及代扣 Then
                               '转帐
                               blnAdd = True
                               For Each objData In objDataTrans
                                   If objData.Key = "K" & lngCardTypeID Then
                                       objData.objBalanceItems.AddItem objItem
                                       objData.objBalanceItems.结算金额 = RoundEx(objData.objBalanceItems.结算金额 + objItem.结算金额, 6)
                                       blnAdd = False
                                       Exit For
                                   End If
                               Next
                               If blnAdd Then  '未找到，需要增加
                                   Set objData = New clsBalanceData
                                   Set objItemsTemp = New clsBalanceItems
                                   
                                   objItemsTemp.AddItem objItem
                                   objItemsTemp.结算金额 = RoundEx(objItemsTemp.结算金额 + objItem.结算金额, 6)
                                   objItemsTemp.收费类型 = 1
                                   objItemsTemp.是否转帐 = True
                                   
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
                                   blnSingleDel = objThirdSwap.zlThirdSwapIsSwapNOCall(lngCardTypeID, bln消费卡, strErrMsg, strExpend)
                                   cllDelSwap.Add Array(lngCardTypeID, IIf(blnSingleDel, 0, 1))
                               End If
                               
                               
                               '加载三方卡
                               objItem.是否退款分交易 = blnSingleDel
                               If blnSingleDel Then
                                    Set objDatasTemp = objDataSingle
                               Else
                                   Set objDatasTemp = objDataMulit
                               End If
                               
                               blnAdd = True
                               For Each objData In objDatasTemp
                                   If objData.Key = "K" & lngCardTypeID Then
                                       objData.objBalanceItems.AddItem objItem
                                       objData.objBalanceItems.结算金额 = RoundEx(objData.objBalanceItems.结算金额 + objItem.结算金额, 6)
                                       blnAdd = False
                                       Exit For
                                   End If
                               Next
                               If blnAdd Then  '未找到，需要增加
                                   Set objData = New clsBalanceData
                                   Set objItemsTemp = New clsBalanceItems
                                   objItemsTemp.AddItem objItem
                                   objItemsTemp.结算金额 = RoundEx(objItemsTemp.结算金额 + objItem.结算金额, 6)
                                   objItemsTemp.收费类型 = 1
                                   objItemsTemp.是否转帐 = False
                                   objData.Key = "K" & lngCardTypeID
                                   Set objData.objBalanceItems = objItemsTemp
                                   objDatasTemp.AddItem objData
                               End If
                               
                           End If
                       End If
                    End If
                    
                ElseIf int结算性质 = 7 Then
                    '老一卡通
                    rsTemp.Filter = "卡类别id=0  And 结算方式='" & str结算方式 & "  And 关联交易ID=" & lng关联交易ID & " And 剩余金额>0 "
                    
                    If rsTemp.EOF Then GoTo GoNext
                    
                    dblTemp = Val(NVL(rsTemp!剩余金额))
                    If dblTemp = 0 Then GoTo GoNext
                    
                    If dblTemp < dblMoney Then dblMoney = dblTemp '本次只能冲剩余金额
                    
                    If dblBalanceSum > dblMoney Then
                        dblBalanceSum = dblBalanceSum - dblMoney
                        dblDelMoney = dblMoney
                    Else
                        dblDelMoney = dblBalanceSum: dblBalanceSum = 0
                    End If
                    
                    If dblDelMoney = 0 Then GoTo GoNext
                    
                    dbl原始金额 = Val(NVL(rsTemp!原始金额))
                    dbl余额 = dbl余额 + Val(NVL(rsTemp!余额))
                    
                    Set objItem = New clsBalanceItem
                    With objItem
                        Set .objCard = zlGetCardFromCardType(lngCardTypeID, bln消费卡, str结算方式)
                        .结算性质 = int结算性质
                        .结算IDs = ""
                        .交易流水号 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("交易流水号")))
                        .交易说明 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("交易说明")))
                        .卡号 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("卡号")))
                        .关联交易ID = lng关联交易ID
                        .结算金额 = RoundEx(-1 * dblDelMoney, 6)
                        .结算方式 = str结算方式
                        .结算号码 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("结算号码")))
                        .结算摘要 = Trim(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("摘要")))
                        .结算类型 = IIf(.结算性质 = 7, 4, IIf(Not bln消费卡, 3, 5)) '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                        .是否预交 = True
                        .是否退款 = True
                        .是否允许编辑 = False
                        .是否允许删除 = True
                        .未退金额 = dbl余额
                        .原始金额 = dbl原始金额
                        .卡类别ID = lngCardTypeID
                        .消费卡 = bln消费卡
                        .是否密文 = IIf(.objCard.卡号密文规则 <> "", True, False)
                        .结帐时间 = CDate(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("收款日期")))
                        .是否允许退现 = .objCard.是否退现
                        .消费卡ID = 0
                        .是否退款分交易 = True
                    End With
                    objOldItems.AddItem objItem
                    objOldItems.收费类型 = 1
                    objOldItems.结算金额 = RoundEx(objOldItems.结算金额 + objItem.结算金额, 6)
    
                Else
                    '其他结算信息
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
    
  
    '1.分交易结算
    For Each objData In objDataSingle
        Set objItems = objData.objBalanceItems
        If objItems.Count <> 0 Then
            For Each objItem In objItems
                objItems_Out.AddItem objItem
                objItems_Out.结算金额 = RoundEx(objItems_Out.结算金额 + objItem.结算金额, 6)
            Next
        End If
    Next
    
    '2.一次交易
     For Each objData In objDataMulit
        Set objItems = objData.objBalanceItems
        If objItems.Count <> 0 Then
            Set objItem = New clsBalanceItem
            With objItem
               Set .objCard = objItems(1).objCard
                .关联交易ID = 0
                .交易流水号 = ""
                .交易说明 = ""
                .结算IDs = objItems(1).结算IDs
                .结算方式 = objItems(1).结算方式
                .结算类型 = objItems(1).结算类型
                .结帐ID = objItems(1).结帐ID
                .结帐时间 = objItems(1).结帐时间
                .结算性质 = objItems(1).结算性质
                .卡类别ID = objItems(1).卡类别ID
                .门诊结帐 = objItems(1).门诊结帐
                .是否保存 = objItems(1).是否保存
                .是否结算 = objItems(1).是否结算
                .是否密文 = objItems(1).是否密文
                .是否强制退现 = objItems(1).是否强制退现
                .是否缺省 = objItems(1).是否缺省
                .是否退款 = objItems(1).是否退款
                .是否退款分交易 = objItems(1).是否退款分交易
                .是否预交 = objItems(1).是否预交
                .是否允许编辑 = objItems(1).是否允许编辑
                .是否允许删除 = objItems(1).是否允许删除
                .是否允许退现 = objItems(1).是否允许退现
                .是否转帐 = objItems(1).是否转帐
                .限制类别 = objItems(1).限制类别
                .消费卡 = objItems(1).消费卡
                .消费卡ID = objItems(1).消费卡ID
                .校对标志 = objItems(1).校对标志
                .预交ID = 0
                .原始金额 = 0
                .缴款金额 = 0
                .结算金额 = 0
                .未退金额 = 0
                Set .objTag = objItems
            
            End With
            For Each objItemTemp In objItems
                objItem.结算金额 = RoundEx(objItem.结算金额 + objItemTemp.结算金额, 6)
                objItem.原始金额 = RoundEx(objItem.原始金额 + objItemTemp.原始金额, 6)
                objItem.未退金额 = RoundEx(objItem.未退金额 + objItemTemp.未退金额, 6)
                objItems.结算金额 = RoundEx(objItems.结算金额 + objItemTemp.结算金额, 6)
            Next
            objItems_Out.AddItem objItem
            objItems_Out.结算金额 = RoundEx(objItems_Out.结算金额 + objItem.结算金额, 6)
        End If
    Next
   '3-转帐
    For Each objData In objDataTrans
      Set objItems = objData.objBalanceItems
      If objItems.Count <> 0 Then
      
        Set objItem = New clsBalanceItem
        With objItem
           Set .objCard = objItems(1).objCard
            .关联交易ID = 0
            .交易流水号 = ""
            .交易说明 = ""
            .结算IDs = objItems(1).结算IDs
            .结算方式 = objItems(1).结算方式
            .结算类型 = objItems(1).结算类型
            .结帐ID = objItems(1).结帐ID
            .结帐时间 = objItems(1).结帐时间
            .结算性质 = objItems(1).结算性质
            .卡类别ID = objItems(1).卡类别ID
            .门诊结帐 = objItems(1).门诊结帐
            .是否保存 = objItems(1).是否保存
            .是否退款分交易 = objItems(1).是否退款分交易
            .是否结算 = objItems(1).是否结算
            .是否密文 = objItems(1).是否密文
            .是否强制退现 = objItems(1).是否强制退现
            .是否缺省 = objItems(1).是否缺省
            .是否退款 = objItems(1).是否退款
            .是否退款分交易 = objItems(1).是否退款分交易
            .是否预交 = objItems(1).是否预交
            .是否允许编辑 = objItems(1).是否允许编辑
            .是否允许删除 = objItems(1).是否允许删除
            .是否允许退现 = objItems(1).是否允许退现
            .是否转帐 = objItems(1).是否转帐
            .限制类别 = objItems(1).限制类别
            .消费卡 = objItems(1).消费卡
            .消费卡ID = objItems(1).消费卡ID
            .校对标志 = objItems(1).校对标志
            .预交ID = 0
            .缴款金额 = 0
            .结算金额 = 0
            .未退金额 = 0
            Set .objTag = objItems
        End With
        For Each objItemTemp In objItems
            objItem.结算金额 = RoundEx(objItem.结算金额 + objItemTemp.结算金额, 6)
            objItem.原始金额 = RoundEx(objItem.原始金额 + objItemTemp.原始金额, 6)
            objItem.未退金额 = RoundEx(objItem.未退金额 + objItemTemp.未退金额, 6)
        Next
        objItems_Out.AddItem objItem
        objItems_Out.结算金额 = RoundEx(objItems_Out.结算金额 + objItem.结算金额, 6)
      End If
    Next
    '3.旧一卡通
    For Each objItem In objOldItems
        objItems_Out.AddItem objItem
        objItems_Out.结算金额 = RoundEx(objItems_Out.结算金额 + objItem.结算金额, 6)
    Next
    '4.普通结算
    For Each objItem In objItemsPt
        objItems_Out.AddItem objItem
        objItems_Out.结算金额 = RoundEx(objItems_Out.结算金额 + objItem.结算金额, 6)
    Next
    zlGetDelDepositItemsFromVsDeposit = True
    
    '释放资源
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


Public Function zlCheckBalancesIsExistFromCardTypeID(ByVal vsBalance As VSFlexGrid, ByVal lng卡类别ID As Long, Optional ByVal bln消费卡 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡类别ID检查是否在结算列表中存在该结算类别的结算信息
    '入参:lng卡类别ID-卡类别ID
    '     bln消费卡-消费卡
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-06-20 10:21:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, blnSquare As Boolean
    Dim i As Long
    On Error GoTo errHandle
    With vsBalance
        For i = 1 To .Rows - 1
            lngCardTypeID = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
            '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            blnSquare = Val(.TextMatrix(i, .ColIndex("类型"))) = 5
            If lngCardTypeID = lng卡类别ID And blnSquare = bln消费卡 Then
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

Public Function zlGet结算方式() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算方式
    '返回:返回结算方式信息集
    '编制:刘兴洪
    '日期:2018-03-29 17:35:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    If Not grs结算方式 Is Nothing Then
        If grs结算方式.State = 1 Then
            grs结算方式.Filter = 0
            Set zlGet结算方式 = grs结算方式: Exit Function
        End If
    End If
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select a.编码,a.名称, a.性质,b.应用场合,nvl(a.应付款,0) as 应付款,nvl(a.应收款,0) as 应收款,nvl(a.缺省标志,0) as 缺省,nvl(b.缺省标志,0) as  场合缺省" & vbNewLine & _
    "   From 结算方式 a, 结算方式应用 b" & vbNewLine & _
    "   Where a.名称 = b.结算方式(+)    "
        
    Set grs结算方式 = zlDatabase.OpenSQLRecord(strSQL, "获取结算方式")
    Set zlGet结算方式 = grs结算方式
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardFromBalanceName(ByVal str结算方式 As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算方式获取卡对象
    '入参:str结算方式-结算方式名称
    '出参:
    '返回:返回卡对象
    '编制:刘兴洪
    '日期:2018-03-30 15:39:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo errHandle
    Set objCard = New Card
    Set rsTemp = zlGet结算方式
    With objCard
        .结算方式 = str结算方式
        .名称 = str结算方式
        rsTemp.Filter = "名称='" & str结算方式 & "'"
        If Not rsTemp.EOF Then
            .结算性质 = Val(NVL(rsTemp!性质))
            .缺省标志 = Val(NVL(rsTemp!场合缺省)) = 1
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
    '功能:根据结算网格中的数据，提取指定行的BalanceItem数据
    '入参:lngRow-指定的行
    '出参:objBalanceItem-
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-30 15:22:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str结算方式 As String, lng卡类别ID As Long, lng消费卡ID As Long
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
        str结算方式 = .TextMatrix(lngRow, .ColIndex("结算方式"))
        If str结算方式 = "" Then Exit Function
        lng卡类别ID = Val(.TextMatrix(lngRow, .ColIndex("卡类别ID")))
        lng消费卡ID = Val(.TextMatrix(lngRow, .ColIndex("消费卡ID")))
        
        If lng卡类别ID = 0 Then
            Set objCard = zlGetCardFromBalanceName(str结算方式)
        Else
            Call gobjSquare.objOneCardComLib.zlGetCard(lng卡类别ID, lng消费卡ID <> 0, objCard)
        End If
        
        varTemp = Split(.TextMatrix(lngRow, .ColIndex("编辑状态")) & "|", "|")
        Set objBalanceItem_Out = New clsBalanceItem
        With objBalanceItem_Out
            Set .objCard = objCard
            .关联交易ID = Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("关联交易ID")))
            
            .交易流水号 = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("交易流水号"))
            .交易说明 = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("交易说明"))
            .结算号码 = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("结算号码"))
            .结算摘要 = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("备注"))
            .卡号 = vsGrid.Cell(flexcpData, lngRow, vsGrid.ColIndex("卡号"))
            .是否密文 = Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("是否密文"))) = 1
            .结算金额 = Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("结算金额")))
            
            .是否允许编辑 = Val(varTemp(0)) = 1
            .是否允许删除 = Val(varTemp(1)) = 1
            .限制类别 = CStr(vsGrid.Cell(flexcpData, lngRow, vsGrid.ColIndex("卡类别ID")))
            .消费卡 = lng消费卡ID <> 0
            .消费卡ID = lng消费卡ID
            .卡类别ID = lng卡类别ID
            .密码 = ""
            .校对标志 = Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("校对标志")))
            .结算性质 = objCard.结算性质
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

Public Function zlGetBalancePatiNums(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal blnZero As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID获取病人结帐的有效住院次数
    '入参:lng病人ID-病人ID
    '     lng主页ID-最后一次的主页ID
    '     blnZero-是否包含零费用
    '出参:
    '返回:返回结帐的有效次数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select Zl_Fun_Getbalancepatinums([1],[2],[3]) As 住院次数 From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人的有效次数", lng病人ID, lng主页ID, IIf(blnZero, 1, 0))
    zlGetBalancePatiNums = NVL(rsTemp!住院次数)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub zlCalcMzDepsitFromMoney(ByVal vsDepositGrid As VSFlexGrid, ByRef dblMoney As Double, ByRef dblToTal_Total As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算门诊冲预交
    '入参:vsDepositGrid-预交网格
    '     dblMoney-本次计算的金额
    '出参:dblToTal_Total-返回冲预交总额
    '编制:刘兴洪
    '日期:2018-12-07 15:36:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    If dblMoney < 0 Then Exit Sub
    With vsDepositGrid
        dblToTal_Total = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" Then
                If Val(.TextMatrix(i, .ColIndex("编辑状态"))) = 0 Then
                    If .TextMatrix(i, .ColIndex("类别")) = "门诊" Then
                          If dblMoney > 0 Then
                              If Val(.TextMatrix(i, .ColIndex("余额"))) <= dblMoney Or dblMoney < 0 Then
                                  .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                              Else
                                  .TextMatrix(i, .ColIndex("冲预交")) = Format(dblMoney, "0.00")
                              End If
                              dblToTal_Total = dblToTal_Total + RoundEx(Val(.TextMatrix(i, .ColIndex("冲预交"))), 2)
                              dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("冲预交")))
                          Else
                             .TextMatrix(i, .ColIndex("冲预交")) = Format(0, "0.00")
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
    ByVal byt门诊预交缺省使用方式 As Byte, ByVal bln中途结帐退预交 As Boolean, _
    Optional ByVal dblMoney As Double = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算冲预交金额
    '入参:bytOperationType-操作类型(0-清除所有冲预交;1-按缺省使用预交款;2-按指定金额来冲预交(按时间先后来分摊）;3-全冲;4-住院全部冲销,不足使用门诊预交
    '     vsDepositGrid-预交网格
    '     objBalanceDatas-当前的结帐数据集
    '     dblMoneny-冲预交金额
    '编制:刘兴洪
    '日期:2018-03-30 11:30:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytCurFun As Byte  '0-全清预交款;1-按结帐金额来冲预交;2-使用所有预交款;3-住院全部冲销,不足使用门诊预交
    Dim dblTotal As Double, i As Long, dblTemp As Double
    Dim bln门诊结帐 As Boolean, bln中途结帐 As Boolean
    Dim bln存在门诊预交 As Boolean
    On Error GoTo errHandle
    
    If objBalanceDatas Is Nothing Then Exit Sub
    
    
    bln门诊结帐 = objBalanceDatas.结算类型 = 1  '是门诊结帐
    bln中途结帐 = objBalanceDatas.是否中途结帐
    
    Select Case bytOperationType
    Case 0  '0-清除所有冲预交
        bytCurFun = 0
    Case 1  '1-按缺省使用预交款
        bytCurFun = 1   '门诊结帐或中途结帐，缺省按结帐金额来使用
        If bln门诊结帐 Then
            Select Case byt门诊预交缺省使用方式   '门诊预交缺省使用方式
            Case 0 ' 0-缺省不使用交;1-按结帐金额使用预交;2-使用所有预交
                bytCurFun = 0
            Case 1 '1-按结帐金额使用预交
                bytCurFun = 1
            Case 2 '2-使用所有预交
                bytCurFun = 2
            End Select
        Else    '住院预交
           If Not bln中途结帐 Or bln中途结帐退预交 Then
                bytCurFun = IIf(gTy_System_Para.TY_Balance.bln允许使用门诊预交, 3, 2)
           End If
        End If
        dblMoney = RoundEx(objBalanceDatas.未付合计, 2)
        
    Case 2 '2-按指定金额来冲预交(按时间先后来分摊）
        bytCurFun = 1
        If dblMoney = 0 Then dblMoney = RoundEx(objBalanceDatas.未付合计, 2)
    Case 3 '3-全冲
        bytCurFun = 2
    Case 4 '4-住院全部冲销,不足使用门诊预交
        bytCurFun = 3
        If dblMoney = 0 Then dblMoney = RoundEx(objBalanceDatas.未付合计, 2)
    Case Else
         bytCurFun = 0
    End Select
    
    If dblMoney < 0 Then dblMoney = 0
    bln存在门诊预交 = False
    With vsDepositGrid
        dblTotal = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" Then
                If Val(.TextMatrix(i, .ColIndex("编辑状态"))) = 0 Then
                    .Cell(flexcpText, i, .ColIndex("冲预交"), i, .ColIndex("冲预交")) = "0.00"
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                    
                    Select Case bytCurFun
                        Case 1 '按结帐金额使用
                            If dblMoney <> 0 Then
                                If Val(.TextMatrix(i, .ColIndex("余额"))) <= dblMoney Then
                                      .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                                Else
                                    .TextMatrix(i, .ColIndex("冲预交")) = Format(dblMoney, "0.00")
                                End If
                                dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("冲预交"))), 2)
                                dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("冲预交")))
                            Else
                               .TextMatrix(i, .ColIndex("冲预交")) = Format(0, "0.00")
                            End If
                        Case 2 '全冲
                            .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                            dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("冲预交"))), 2)
                        Case 3 '住院全部冲销,不足使用门诊预交
                            If .TextMatrix(i, .ColIndex("类别")) <> "门诊" Then
                                .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                                dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("冲预交"))), 2)
                                dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("冲预交")))
                            Else
                                .Cell(flexcpText, i, .ColIndex("冲预交"), i, .ColIndex("冲预交")) = "0.00"
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                                If bln存在门诊预交 = False Then bln存在门诊预交 = True
                            End If
                        Case 0 '清除
                            .TextMatrix(i, .ColIndex("冲预交")) = Format(0, "0.00")
                        Case Else
                    End Select
                Else
                    dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("冲预交"))), 2)
                    dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("冲预交")))
                End If
            End If
            Next
    End With
    dblTemp = 0
    If bytCurFun = 3 And bln存在门诊预交 Then Call zlCalcMzDepsitFromMoney(vsDepositGrid, dblMoney, dblTemp)
    objBalanceDatas.冲预交合计 = RoundEx(dblTotal + dblTemp, 6)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlLoadDepositListFromRecord(ByVal bytRecalDepsoit As Byte, ByVal rsDeposit As ADODB.Recordset, ByVal dbl未付合计 As Double, vsDeposit As VSFlexGrid, _
    ByRef dblTotal_Out As Double, ByRef intCountBill_Out As Integer, ByVal lngModul As Long, _
    Optional strFormName As String, Optional strRegKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据预交记录集将预交信息加载到预交列表中
    '入参:rsDeposit-预交记录集信息
    '    dbl未付合计-未付合计
    '    bytRecalDepsoit-冲预交计算状态:0-不计算,1-按金额计算;2- 全部冲销;3-住院全部冲销,不足使用门诊预交
    '出参:dblTotal_Out-冲预交总计
    '     intCountBill_Out-涉及的票据张数
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-25 14:29:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lng卡类别ID As Long, dblTotal As Double
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
            
            .RowData(i) = Val(NVL(rsDeposit!记录状态))
             lng卡类别ID = Val(NVL(rsDeposit!卡类别ID))
             If lng卡类别ID = 0 Then lng卡类别ID = Val(NVL(rsDeposit!结算卡序号))
            
            .TextMatrix(i, .ColIndex("ID")) = rsDeposit!ID
            .TextMatrix(i, .ColIndex("单据号")) = rsDeposit!NO
            .TextMatrix(i, .ColIndex("类别")) = NVL(rsDeposit!预交类别)
            
            .TextMatrix(i, .ColIndex("票据号")) = "" & rsDeposit!票据号
            .TextMatrix(i, .ColIndex("收款日期")) = Format(rsDeposit!日期, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("结算方式")) = NVL(rsDeposit!结算方式)
            .TextMatrix(i, .ColIndex("余额")) = Format(rsDeposit!金额, "0.00")
            .TextMatrix(i, .ColIndex("预交ID")) = NVL(rsDeposit!预交ID)
            .TextMatrix(i, .ColIndex("关联交易ID")) = NVL(rsDeposit!关联交易ID)
            .TextMatrix(i, .ColIndex("卡类别ID")) = lng卡类别ID
            .TextMatrix(i, .ColIndex("是否消费卡")) = Val(NVL(rsDeposit!是否消费卡))
            .TextMatrix(i, .ColIndex("卡号")) = NVL(rsDeposit!卡号)
            .TextMatrix(i, .ColIndex("卡类别名称")) = NVL(rsDeposit!卡类别名称)
            .TextMatrix(i, .ColIndex("交易流水号")) = NVL(rsDeposit!交易流水号)
            .TextMatrix(i, .ColIndex("交易说明")) = NVL(rsDeposit!交易说明)
            .TextMatrix(i, .ColIndex("是否退现")) = Val(NVL(rsDeposit!是否退现))
            .TextMatrix(i, .ColIndex("是否全退")) = Val(NVL(rsDeposit!是否全退))
            .TextMatrix(i, .ColIndex("是否缺省退现")) = Val(NVL(rsDeposit!是否缺省退现))
            .TextMatrix(i, .ColIndex("是否转帐及代扣")) = Val(NVL(rsDeposit!转帐及代扣))
            .TextMatrix(i, .ColIndex("结算性质")) = Val(NVL(rsDeposit!结算性质))
            .TextMatrix(i, .ColIndex("原始金额")) = Val(NVL(rsDeposit!原始金额))
            .TextMatrix(i, .ColIndex("结算号码")) = NVL(rsDeposit!结算号码)
            .TextMatrix(i, .ColIndex("摘要")) = NVL(rsDeposit!摘要)
            
            Select Case bytRecalDepsoit
            Case 1 '按金额冲销
                If Val(NVL(rsDeposit!金额)) <= dbl未付合计 Then
                    .TextMatrix(i, .ColIndex("冲预交")) = Format(rsDeposit!金额, "0.00")
                    dbl未付合计 = dbl未付合计 - RoundEx(Val(NVL(rsDeposit!金额)), 2)
                ElseIf dbl未付合计 <> 0 Then
                    .TextMatrix(i, .ColIndex("冲预交")) = Format(dbl未付合计, "0.00")
                    dbl未付合计 = 0
                End If
            Case 2 '全部冲销
                .TextMatrix(i, .ColIndex("冲预交")) = Format(rsDeposit!金额, "0.00")
            Case 3 '3-住院全部冲销,不足使用门诊预交
                If .TextMatrix(i, .ColIndex("类别")) <> "门诊" Then
                    .TextMatrix(i, .ColIndex("冲预交")) = Format(rsDeposit!金额, "0.00")
                    dbl未付合计 = dbl未付合计 - RoundEx(Val(NVL(rsDeposit!金额)), 2)
                Else
                    If Val(NVL(rsDeposit!金额)) <= dbl未付合计 Then
                        .TextMatrix(i, .ColIndex("冲预交")) = Format(rsDeposit!金额, "0.00")
                        dbl未付合计 = dbl未付合计 - RoundEx(Val(NVL(rsDeposit!金额)), 2)
                    ElseIf dbl未付合计 > 0 Then
                        .TextMatrix(i, .ColIndex("冲预交")) = Format(dbl未付合计, "0.00")
                        dbl未付合计 = 0
                    End If
                End If
            Case Else '0 -不计算
            End Select
            
            dblTotal = dblTotal + RoundEx(Val(NVL(rsDeposit!金额)), 2)
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

Public Function zlGetClassMoney(ByVal lng结帐ID As Long, ByVal dbl当前结帐金额 As Double, ByRef rsMoney As ADODB.Recordset, _
    rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '入参:lng结帐ID-结帐ID,为0时,以RsFeeList为准
    '     dbl当前结帐金额-当前结帐金额(主要是分摊时使用)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, dblMoney As Double
    Dim dblTemp As Double
    
    On Error GoTo errHandle
    
    '初始化数据结构
    Set rsMoney = New ADODB.Recordset
    rsMoney.Fields.Append "收费类别", adVarChar, 10, adFldIsNullable
    rsMoney.Fields.Append "金额", adDouble, , adFldIsNullable
    rsMoney.CursorLocation = adUseClient
    rsMoney.LockType = adLockOptimistic
    rsMoney.CursorType = adOpenStatic
    rsMoney.Open
        
    If lng结帐ID <> 0 Then
        strSQL = "" & _
        "   Select  A.收费类别,nvl(sum(A.结帐金额) ,0) as 金额   " & _
        "   From 门诊费用记录 A" & _
        "   Where A.结帐ID=[1] Group by A.收费类别 " & _
        "   Union ALL " & _
        "   Select  A.收费类别,nvl(sum(A.结帐金额) ,0) as 金额   " & _
        "   From 住院费用记录 A" & _
        "   Where A.结帐ID=[1] Group by A.收费类别 "
        strSQL = "Select 收费类别,Sum(金额) as 金额 From (" & strSQL & ")  Group by  收费类别"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取收费用类别", lng结帐ID)
    
        With rsTemp
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                rsMoney.Find "收费类别='" & NVL(!收费类别, "无") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!收费类别 = NVL(!收费类别, "无")
                rsMoney!金额 = Val(NVL(rsMoney!金额)) + Val(NVL(!金额))
                rsMoney.Update
                .MoveNext
            Loop
        End With
        zlGetClassMoney = True
        Exit Function
    End If
    
    If rsFeeList Is Nothing Then Exit Function
    
    With rsFeeList
        dblMoney = dbl当前结帐金额
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dblTemp = Val(NVL(!未结金额))
            If RoundEx(dblMoney - dblTemp, gbytDec) <= 0 Then
                dblTemp = dblMoney
            End If
            If dblTemp <> 0 And dblMoney <> 0 Then
                rsMoney.Find "收费类别='" & NVL(!收费类别, "无") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!收费类别 = NVL(!收费类别, "无")
                rsMoney!金额 = Val(NVL(rsMoney!金额)) + dblTemp
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
    '功能:根据卡对象，从结算列表中获取相关的结算数据
    '入参:objCard-当前卡对象
    '出参:objBalanceItems_Out-返回交易数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-30 15:22:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str结算方式 As String, lng卡类别ID As Long, lng消费卡ID As Long
    Dim varTemp As Variant
    Dim objItem As clsBalanceItem
    
    
    On Error GoTo errHandle
    
    If objCard Is Nothing Then Exit Function
    Set objBalanceItems_out = New clsBalanceItems
    
    With vsGrid
        For i = 1 To .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, i, objItem) Then
                If (objItem.卡类别ID = objCard.接口序号 And objItem.消费卡 = objCard.消费卡 And (objItem.消费卡ID <> 0 And objItem.消费卡 Or objItem.消费卡 = False)) Or (objCard.接口序号 <= 0 And objCard.结算方式 = objItem.结算方式) Then
                    objBalanceItems_out.AddItem objItem
                    objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
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

Public Function zlGetBalanceItemsFromGrid(ByVal vsGrid As VSFlexGrid, ByVal int类型 As Integer, ByRef objBalanceItems_out As clsBalanceItems, Optional objBalanceItem As clsBalanceItem, Optional bln退款 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据网格，获取所有的消费卡结算信息集
    '入参:vsGrid-网格对象
    '     int类型-0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    '     objItem-如果Nothing,则表示按类型取数,否则按当前项目取数
    '     bln退款-是否获取退款：bln退款-true,返回:objitem结算金额的相返数，主要是方便退款)
    '出参:objBalanceItems_Out-返回消费卡相关信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-11 18:42:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    Dim i As Long
    On Error GoTo errHandle
    Set objBalanceItems_out = New clsBalanceItems
    With vsGrid
        For i = 1 To .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, i, objItem) Then
                If Not objBalanceItem Is Nothing Then
                    If objItem.结算类型 = int类型 And ( _
                        (objBalanceItem.关联交易ID = objItem.关联交易ID And objBalanceItem.卡类别ID = objItem.卡类别ID And objBalanceItem.消费卡 = objItem.消费卡) _
                        Or (objBalanceItem.卡类别ID <= 0 And objItem.卡类别ID <= 0 And objBalanceItem.结算方式 = objItem.结算方式)) Then
                        
                        Set objItem = zlCopyNewItemFromBalanceItem(objItem)
                        If bln退款 Then
                            If objItem.原始金额 = 0 Then objItem.原始金额 = objItem.结算金额
                            If objItem.未退金额 = 0 Then objItem.未退金额 = objItem.结算金额
                            objItem.结算金额 = RoundEx(-1 * objItem.结算金额, 6)
                        End If
                        
                        objBalanceItems_out.AddItem objItem
                        objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
                    End If
                Else
                    If objItem.结算类型 = int类型 Then
                        Set objItem = zlCopyNewItemFromBalanceItem(objItem)
                        If bln退款 Then
                            If objItem.原始金额 = 0 Then objItem.原始金额 = objItem.结算金额
                            If objItem.未退金额 = 0 Then objItem.未退金额 = objItem.结算金额
                            objItem.结算金额 = RoundEx(-1 * objItem.结算金额, 6)
                        End If
                        objBalanceItems_out.AddItem objItem
                        objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
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
    '功能:获取退三方缴预交的退款信息
    '入参:
    '出参:clsBalanceDatas-返回结算信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-16 14:13:39
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
                If objItem.是否预交 And objItem.是否退款 And objItem.是否结算 = False And objItem.卡类别ID <> 0 And objItem.消费卡 = False Then
                    If InStr(strCardTypeIDs & ",", "," & objItem.卡类别ID & ",") = 0 Then
                        Set objItems = New clsBalanceItems
                        Set objData = New clsBalanceData
                        Set objData.objBalanceItems = objItems
                        objData.Key = "K" & objItem.卡类别ID
                        Call objBalanceDatas_Out.AddItem(objData, "K" & objItem.卡类别ID)
                        strCardTypeIDs = strCardTypeIDs & "," & objItem.卡类别ID
                    End If
                    
                    objBalanceDatas_Out("K" & objItem.卡类别ID).objBalanceItems.AddItem objItem
                    With objBalanceDatas_Out("K" & objItem.卡类别ID).objBalanceItems
                        .结算金额 = .结算金额 + objItem.结算金额
                        .收费类型 = 1
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

Public Function zlThirdDelMoneyIsExistFromVsGrid(ByVal vsGrid As VSFlexGrid, ByVal lng卡类别ID As Long, ByVal lng关联交易ID As Long, str交易流水号 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定的卡类别及交易是否在结算方式信息中存在退款
    '入参:
    '出参:
    '返回:存在退款返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-18 23:42:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    Dim i As Long
    
    On Error GoTo errHandle
    
    With vsGrid
        For i = 1 To .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, i, objItem) Then
               If objItem.卡类别ID = lng卡类别ID And (objItem.关联交易ID = lng关联交易ID Or objItem.objCard.是否转帐及代扣) And objItem.结算金额 < 0 Then
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
    '功能:复制一个新的Item
    '入参:objOldItem-旧的Item对象
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-19 14:14:23
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
        .关联交易ID = objOldItem.关联交易ID
        .交易流水号 = objOldItem.交易流水号
        .交易说明 = objOldItem.交易说明
        .缴款金额 = objOldItem.缴款金额
        .结算IDs = objOldItem.结算IDs
        
        .结算方式 = objOldItem.结算方式
        .结算号码 = objOldItem.结算号码
        .结算金额 = objOldItem.结算金额
        .结算类型 = objOldItem.结算类型
        .结算性质 = objOldItem.结算性质
        .结算摘要 = objOldItem.结算摘要
        .结帐ID = objOldItem.结帐ID
        .结帐时间 = objOldItem.结帐时间
        .冲销ID = objOldItem.冲销ID
        
        .卡号 = objOldItem.卡号
        .交易流水号 = objOldItem.交易流水号
        .交易说明 = objOldItem.交易说明
        .卡类别ID = objOldItem.卡类别ID
        .密码 = objOldItem.密码
        .是否结算 = objOldItem.是否结算
        .是否密文 = objOldItem.是否密文
        .是否缺省 = objOldItem.是否缺省
        .是否退款 = objOldItem.是否退款
        .是否预交 = objOldItem.是否预交
        .是否允许编辑 = objOldItem.是否允许编辑
        .是否允许删除 = objOldItem.是否允许删除
        .是否允许退现 = objOldItem.是否允许退现
        .是否保存 = objOldItem.是否保存
        .是否退款分交易 = objOldItem.是否退款分交易
        .未退金额 = objOldItem.未退金额
        .误差费 = objOldItem.误差费
        .限制类别 = objOldItem.限制类别
        .消费卡 = objOldItem.消费卡
        .消费卡ID = objOldItem.消费卡ID
        .校对标志 = objOldItem.校对标志
        .原始金额 = objOldItem.原始金额
        .帐户余额 = objOldItem.帐户余额
        .退款交易流水号 = objOldItem.退款交易流水号
        .退款交易说明 = objOldItem.退款交易说明
        .找补 = objOldItem.找补
        .行号 = objOldItem.行号
        .预交ID = objOldItem.预交ID
        .是否脱机医保 = objOldItem.是否脱机医保
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
    '功能:根据一个卡对象，复制为新的卡对象
    '入参:objOldCard-旧卡
    '返回:返回新的Card对象
    '编制:刘兴洪
    '日期:2018-04-19 14:25:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Set objCard = New Card
    If objOldCard Is Nothing Then Set zlCopyNewCardFromCard = Nothing: Exit Function
    
    On Error GoTo errHandle
    With objOldCard
        objCard.备注 = .备注
        objCard.短名 = .短名
        objCard.功能键 = .功能键
        objCard.接口编码 = .接口编码
        objCard.接口程序名 = .接口程序名
        objCard.接口序号 = .接口序号
        objCard.结算方式 = .结算方式
        objCard.结算性质 = .结算性质
        objCard.卡号密文规则 = .卡号密文规则
        objCard.卡号长度 = .卡号长度
        objCard.卡号重复使用 = .卡号重复使用
        objCard.可否设置 = .可否设置
        objCard.快键 = .快键
        objCard.密码规则 = .密码规则
        objCard.密码输入限制 = .密码输入限制
        objCard.密码长度 = .密码长度
        objCard.密码长度限制 = .密码长度限制
        objCard.名称 = .名称
        objCard.模糊查找项 = .模糊查找项
        objCard.启用 = .启用
        objCard.前缀文本 = .前缀文本
        objCard.缺省标志 = .缺省标志
        objCard.设备是否启用回车 = .设备是否启用回车
        objCard.是否持卡消费 = .是否持卡消费
        objCard.是否存在帐户 = .是否存在帐户
        objCard.是否发卡 = .是否发卡
        objCard.是否非接触式读卡 = .是否非接触式读卡
        objCard.是否接触式读卡 = .是否接触式读卡
        objCard.是否模糊查找 = .是否模糊查找
        objCard.是否全退 = .是否全退
        objCard.是否缺省密码 = .是否缺省密码
        objCard.是否扫描 = .是否扫描
        objCard.是否刷卡 = .是否刷卡
        objCard.是否退款验卡 = .是否退款验卡
        objCard.是否退现 = .是否退现
        objCard.是否写卡 = .是否写卡
        objCard.是否严格控制 = .是否严格控制
        objCard.是否证件 = .是否证件
        objCard.是否制卡 = .是否制卡
        objCard.是否转帐及代扣 = .是否转帐及代扣
        objCard.是否自动读取 = .是否自动读取
        objCard.特定项目 = .特定项目
        objCard.图像标识 = .图像标识
        objCard.系统 = .系统
        objCard.消费卡 = .消费卡
        objCard.支付启用 = .支付启用
        objCard.支付图像标识 = .支付图像标识
        objCard.自动读取间隔 = .自动读取间隔
        objCard.自制卡 = .自制卡
        objCard.是否支持扫码付 = .是否支持扫码付
        objCard.是否独立结算 = .是否独立结算
    End With
    Set zlCopyNewCardFromCard = objCard
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    zlCopyNewCardFromCard = objCard
End Function

Public Function zlGetCardFromCardType(ByVal lng卡类别ID As Long, bln消费卡 As Boolean, ByVal str结算方式 As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡类别ID获取卡对象
    '入参:lng卡类别ID-卡类别ID
    '     bln消费卡-是否消费卡
    '     str结算方式-结算方式
    '出参:
    '返回:成功卡对象
    '编制:刘兴洪
    '日期:2018-04-02 14:29:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As New Card
    On Error GoTo errHandle
    If lng卡类别ID <> 0 Then
        'zlGetCard(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean,  ByRef objCard As Card) As Boolean
        If gobjSquare.objOneCardComLib.zlGetCard(lng卡类别ID, bln消费卡, objCard) = False Then
            Set objCard = zlGetCardFromBalanceName(str结算方式)
        End If
    Else
        Set objCard = zlGetCardFromBalanceName(str结算方式)
    End If
    Set zlGetCardFromCardType = objCard: Exit Function

    zlGetCardFromCardType = True
    Exit Function
errHandle:
    Set objCard = zlGetCardFromBalanceName(str结算方式)
    Set zlGetCardFromCardType = objCard: Exit Function
End Function

Public Sub zlClearBalanceFromBalanceGrid(ByRef vsGrid As VSFlexGrid, ByVal strBalance As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算方式来清除结算信息行
    '编制:刘兴洪
    '日期:2018-04-16 11:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, objItem As clsBalanceItem
    On Error GoTo errHandle
    With vsGrid
        j = 1
        Do While j <= .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, j, objItem) Then
                If objItem.结算方式 = strBalance And objItem.结算类型 = 0 Then
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
    '功能:清除退预交款的所有结算信息
    '入参:objItems-本次普通结算信息集(即同步删除结算类型为0的结算记录)
    '编制:刘兴洪
    '日期:2018-04-16 11:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, objItem As clsBalanceItem
    Dim blnDel As Boolean, objItemTemp As clsBalanceItem
    
    On Error GoTo errHandle
    With vsGrid
        j = 1
        Do While j <= .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, j, objItem) Then
            
                blnDel = (objItem.是否预交 Or objItem.结算性质 = 9 Or objItem.Tag = "指定预交退款")
                
                If Not objItems Is Nothing And Not blnDel Then
                    '需要排除预交款重新计算时，退指定结算方式后，再次计算时未返回指定结算方式。所以也要一并清除
                    For Each objItemTemp In objItems
                        If objItemTemp.结算类型 = 0 And objItem.结算方式 = objItemTemp.结算方式 And objItemTemp.结算类型 = objItem.结算类型 Then
                            '找到了:
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
    '功能:获取普通的结算信息项
    '入参:vsBalance-网格控件
    '出参:objPtItems_Out-普通结算信息项
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-29 12:10:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As clsBalanceItem
    
    On Error GoTo errHandle
    Set objPtItems_Out = New clsBalanceItems
    With vsBalance
        For i = 1 To .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsBalance, i, objItem) Then
                '需要排除预交款重新计算时，退指定结算方式后，再次计算时未返回指定结算方式。所以也要一并清除
                If objItem.结算类型 = 0 And objItem.结算性质 <> 9 Then
                    objPtItems_Out.AddItem objItem
                    objPtItems_Out.结算金额 = RoundEx(objPtItems_Out.结算金额 + objItem.结算金额, 6)
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
    '功能:获取普通的结算信息项
    '入参:vsBalance-网格控件
    '出参:objPtItems_Out-普通结算信息项
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-29 12:10:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim j As Long, objItem As clsBalanceItem
   Dim blnDel As Boolean, objItemTemp As clsBalanceItem

    
    On Error GoTo errHandle
      
    If objCurItems Is Nothing Then Exit Function
    With vsBalance
        j = 1
        Do While j <= .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsBalance, j, objItem) Then
                '需要排除预交款重新计算时，退指定结算方式后，再次计算时未返回指定结算方式。所以也要一并清除
                blnDel = False
                For Each objItemTemp In objCurItems
                    If objItem.卡类别ID = objItemTemp.卡类别ID And objItemTemp.结算类型 = objItem.结算类型 And objItemTemp.消费卡 = objItem.消费卡 And objItemTemp.关联交易ID = objItem.关联交易ID Then
                        '找到了:
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
    '功能:清除退预交款的所有结算信息(只清除普通结算信息)
    '入参:objItems-本次普通结算信息集(即同步删除结算类型为0的结算记录)
    '编制:刘兴洪
    '日期:2018-04-16 11:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, objItem As clsBalanceItem
    Dim blnDel As Boolean, objItemTemp As clsBalanceItem
    
    On Error GoTo errHandle
    If objItems Is Nothing Then Exit Sub
    
    With vsGrid
        j = 1
        Do While j <= .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsGrid, j, objItem) Then
             
                '需要排除预交款重新计算时，退指定结算方式后，再次计算时未返回指定结算方式。所以也要一并清除
                blnDel = False
                For Each objItemTemp In objItems
                    If objItemTemp.结算类型 = 0 And objItem.结算方式 = objItemTemp.结算方式 And objItemTemp.结算类型 = objItem.结算类型 Then
                        '找到了:
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

Public Function zlGetThirdDelRecordFromBalanceID(ByVal lng结帐ID As Long, ByRef rsThirdDel_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID,获取三方退款信息集
    '入参:lng结帐ID-结帐ID
    '出参:rsThirdDel_Out-返回三方退款信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-17 15:45:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    
    strSQL = " " & _
    "   Select a.结帐id, a.记录id As 预交id,nvl(a.卡类别ID,b.卡类别ID ) as 卡类别ID,a.卡号, a.金额,b.交易流水号, b.交易说明, a.是否未退, a.是否转帐, b.关联交易id,b.结算号码,B.摘要,b.金额 as 原始金额" & vbCrLf & _
    "   From 三方退款信息 A, 病人预交记录 B " & vbCrLf & _
    "   Where a.结帐id =[1] And a.记录id = b.Id" & vbCrLf & _
    " "
    Set rsThirdDel_Out = zlDatabase.OpenSQLRecord(strSQL, "根据结帐ID获取三方退款信息明细", lng结帐ID)
    
    zlGetThirdDelRecordFromBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetDelBalanceItemsFromRecord(ByVal objCurItem As clsBalanceItem, _
    ByVal rsThirdDelRecord As ADODB.Recordset, ByRef objDelItems_Out As clsBalanceItems, _
    Optional ByVal bln退款分交易 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据退款记录，获取三方退款信息明细集
    '入参:objCurItem-当前的结算信息
    '     rsThirdDelRecord-三方退款信息集
    '     bln退款分交易-退款是否分交易
    '出参:objDelItems_out-返回三方退款信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-17 16:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    Dim objCard As Card
    On Error GoTo errHandle
    Set objDelItems_Out = New clsBalanceItems
    If rsThirdDelRecord Is Nothing Then Exit Function
    If rsThirdDelRecord.State <> 1 Then Exit Function
    If bln退款分交易 Then
        rsThirdDelRecord.Filter = "卡类别ID=" & objCurItem.卡类别ID & " And 预交ID=" & objCurItem.预交ID
    Else
        rsThirdDelRecord.Filter = "卡类别ID=" & objCurItem.卡类别ID
    End If
    With rsThirdDelRecord
        Set objCard = zlGetCardFromCardType(objCurItem.卡类别ID, False, "")
        Do While Not .EOF
            Set objItem = New clsBalanceItem
            Set objItem.objCard = objCard
            objItem.关联交易ID = Val(NVL(!关联交易ID))
            objItem.预交ID = Val(NVL(!预交ID))
            objItem.校对标志 = IIf(Val(NVL(!是否未退)) = 1, 1, 2)
            objItem.行号 = objCurItem.行号
            objItem.交易流水号 = Trim(NVL(!交易流水号))
            objItem.交易说明 = Trim(NVL(!交易说明))
            objItem.缴款金额 = 0
            objItem.结算IDs = objCurItem.结算IDs
            objItem.结算方式 = objCurItem.结算方式
            objItem.结算号码 = Trim(NVL(!结算号码))
            objItem.结算金额 = RoundEx(-1 * Val(Trim(NVL(!金额))), 6)
            objItem.结算类型 = objCurItem.结算类型
            objItem.结算性质 = objCurItem.结算性质
            objItem.结算摘要 = Trim(NVL(!摘要))
            objItem.结帐ID = objCurItem.结帐ID
            objItem.结帐时间 = objCurItem.结帐时间
            objItem.卡号 = Trim(NVL(!卡号))
            objItem.卡类别ID = objCurItem.卡类别ID
            objItem.门诊结帐 = objCurItem.门诊结帐
            objItem.密码 = objCurItem.密码
            objItem.剩余金额 = 0
            objItem.是否保存 = objCurItem.是否保存
            objItem.是否结算 = objCurItem.校对标志 = 2
            objItem.是否密文 = objCurItem.是否密文
            objItem.是否强制退现 = objCurItem.是否强制退现
            objItem.是否缺省 = objCurItem.是否缺省
            objItem.是否退款 = objCurItem.是否退款
            objItem.是否退款分交易 = objCurItem.是否退款分交易
            objItem.是否预交 = objCurItem.是否预交
            objItem.是否允许编辑 = objCurItem.是否允许编辑
            objItem.是否允许删除 = objCurItem.是否允许删除
            objItem.是否允许退现 = objCurItem.是否允许退现
            objItem.是否转帐 = objCard.是否转帐及代扣
            If Val(NVL(!是否转帐)) = 1 Then objItem.是否转帐 = True: objItem.objCard.是否转帐及代扣 = True
            objItem.未退金额 = 0
            objItem.误差费 = 0
            objItem.限制类别 = objCurItem.限制类别
            objItem.消费卡 = objCurItem.消费卡
            objItem.消费卡ID = objCurItem.消费卡ID
            objItem.原始金额 = Val(Trim(NVL(!原始金额)))
            objItem.帐户余额 = 0
            objItem.找补 = 0
            If objItem.是否转帐 Then objCurItem.是否转帐 = True
            objDelItems_Out.AddItem objItem
            objDelItems_Out.结算金额 = objDelItems_Out.结算金额 + objItem.结算金额
            objDelItems_Out.类型 = objItem.结算类型
            objDelItems_Out.是否转帐 = objItem.是否转帐
            objDelItems_Out.退费结帐IDs = objCurItem.结算IDs
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
Public Function zlGetBalanceItemsFromRecord(ByVal lng病人ID As Long, ByVal bytMCMode As Byte, intInsure As Integer, ByVal objThirdSwap As clsThirdSwap, _
    ByVal bln作废 As Boolean, ByVal rsBalanceRecord As ADODB.Recordset, ByRef objBalanceInfor As clsBalanceInfo, ByRef objBalanceItems_out As clsBalanceItems, _
    Optional strErrMsg_out As String, Optional bytCurType As Byte, Optional ByRef bln医保作废作退_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐记录数据返回指定的结算数据集
    '入参:objCard-卡对象
    '     lng关联交易Id-关联交易ID
    '     bytCurType-当前操作的类型:0-结帐;1-结帐作废;2-异常重结;3-异常重退,4-正常单据查看;5-作废单据查看;6-异常结算作废
    '出参:objBalanceItems_Out-返结算数据
    '     objBalanceInfor-当前结帐信息
    '     strErrMsg_Out-返回错误信息
    '     bln医保作废作退_Out-医保作废是否全退
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-30 10:31:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPtItems As clsBalanceItems, objItems As clsBalanceItems, objItemsTemp As clsBalanceItems
    Dim objMulitItems As clsBalanceItems, objTransItems As clsBalanceItems
    Dim objItemTemp As clsBalanceItem, objItem As clsBalanceItem
    Dim bln转帐 As Boolean, blnSingleDel As Boolean, strErrMsg As String, strExpend As String
    Dim cllDelSwap As Collection, blnNoBalanceData As Boolean '是否有病人预交记录
    Dim strTemp As String, strDefaultBalance As String
    Dim dblMoney As Double, lng卡类别ID As Long
    Dim blnAdd As Boolean, blnDelCash As Boolean, blnFind As Boolean
    Dim rsThirdDel As ADODB.Recordset, i As Long
    Dim objCard As Card, bln是否保存 As Boolean
    Dim intSign As Integer  '正负数
    Dim strCardTypes As String, rsThirdDelClone As ADODB.Recordset
    Dim strMulitCardTypeIds As String '多笔交易合并的卡类别IDs:卡类别ID,...
    Dim strSingleCardTypeIds As String  '分交易的卡类别IDs:卡类别ID|预交ID,...
    Dim blnReturnCash As Boolean
    
    On Error GoTo errHandle
    
    bln医保作废作退_Out = True
    
    If Not (bytCurType = 4 Or bytCurType = 5) Then '不为查看时才处理
        If zlGetThirdDelRecordFromBalanceID(objBalanceInfor.结帐ID, rsThirdDel) = False Then Exit Function
    End If
    
    If objBalanceInfor Is Nothing Then Set objBalanceInfor = New clsBalanceInfo
    Set objBalanceItems_out = New clsBalanceItems
    
    objBalanceInfor.当前结帐 = 0
    objBalanceInfor.已付合计 = 0
    objBalanceInfor.医保支付合计 = 0
    
    
    strTemp = "  未找到原始的结算记录"
    If rsBalanceRecord Is Nothing Then strErrMsg_out = strTemp: Exit Function
    If rsBalanceRecord.State <> 1 Then strErrMsg_out = strTemp: Exit Function
    
    bln是否保存 = IIf(bytCurType = 1 Or bytCurType = 4 Or bytCurType = 5, False, True) '作废时代表未进行冲销保存.
    
    Set objPtItems = New clsBalanceItems
    
    rsBalanceRecord.Sort = "类型,卡类别ID,关联交易ID"
    intSign = IIf(bytCurType = 3 Or bytCurType = 5, -1, 1)
    
    
     Set objItems = Nothing
    'objBalanceInfor.冲预交合计 = 0
    With rsBalanceRecord
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            objBalanceInfor.当前结帐 = objBalanceInfor.当前结帐 + intSign * Val(NVL(!冲预交))
            
            lng卡类别ID = Val(NVL(!卡类别ID))
            Select Case Val(NVL(!类型))
            Case 0  '普通结算
                If NVL(!结算方式) <> "" Then
                    Set objItem = New clsBalanceItem
                    Set objCard = zlGetCardFromCardType(0, False, NVL(!结算方式))
                    If Not zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) Then Exit Function
                    objItem.结帐ID = objBalanceInfor.结帐ID
                    objItem.结算IDs = objBalanceInfor.结帐ID
                    objItem.冲销ID = objBalanceInfor.冲销ID
                    objItem.结帐时间 = objBalanceInfor.结帐时间
                    objItem.结算类型 = Val(NVL(!类型))
                    objItem.结算金额 = RoundEx(intSign * objItem.结算金额, 6)
                     
                    objItem.是否允许删除 = True
                    objItem.是否允许退现 = True
                    objItem.是否允许编辑 = False
                    
                    
                    If Val(NVL(!性质)) = 1 And (bytCurType <> 4 And bytCurType <> 5) Then '现金特殊处理
                        objBalanceInfor.现金支付 = Val(NVL(!冲预交))
                    Else
                        objBalanceItems_out.AddItem objItem
                        objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + objItem.结算金额, 6)
                        objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
                    End If
                End If
            Case 1 '预交款
                'objBalanceInfor.冲预交合计 = objBalanceInfor.冲预交合计 + Val(nvl(!冲预交))
                'objBalanceInfor.是否保存预交 = True
            Case 2 '医保
                Set objCard = zlGetCardFromCardType(lng卡类别ID, Val(NVL(!类型)) = 5, NVL(!结算方式))
                If zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) = False Then Exit Function
                blnAdd = True
                If bln作废 Then
                       Select Case Val(NVL(!性质))
                       Case 3   '个人帐户
                            If bytMCMode = 1 And Not objThirdSwap.zlGetYbPara.门诊病人结算作废 Then
                                '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                                blnAdd = False '退现
                            Else
                                blnAdd = gclsInsure.GetCapability(IIf(bytMCMode = 1, support门诊结算作废, support住院结算作废), lng病人ID, intInsure, NVL(!结算方式))
                            End If
                       Case 4  '医保基金
                            If bytMCMode = 1 And Not objThirdSwap.zlGetYbPara.门诊病人结算作废 Then
                                blnAdd = True '原样退回
                            Else
                                blnAdd = gclsInsure.GetCapability(IIf(bytMCMode = 1, support门诊结算作废, support住院结算作废), lng病人ID, intInsure, NVL(!结算方式))  '检查是否支持原样退
                            End If
                       End Select
                End If
                
                If blnAdd = False Then bln医保作废作退_Out = False
                
                If blnAdd Then
                    objItem.结算类型 = Val(NVL(!类型))
                    objItem.结帐ID = objBalanceInfor.结帐ID
                    objItem.结算IDs = objBalanceInfor.结帐ID
                    objItem.结帐时间 = objBalanceInfor.结帐时间
                    objItem.冲销ID = objBalanceInfor.冲销ID
                    objItem.是否允许删除 = False
                    objItem.是否允许退现 = False
                    objItem.是否允许编辑 = False
                    objItem.是否保存 = bln是否保存
                    objItem.结算金额 = RoundEx(intSign * objItem.结算金额, 6)
                    objBalanceItems_out.AddItem objItem
                    
                    objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + objItem.结算金额, 6)
                    objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 5)
                    If objItem.校对标志 <> 1 Then
                        objBalanceInfor.医保支付合计 = RoundEx(objBalanceInfor.医保支付合计 + objItem.结算金额, 5)
                    End If
                End If
            Case 3 '一卡通
                
                If lng卡类别ID = 0 Then strErrMsg_out = "结算数据有误，三方结算(" & NVL(!结算方式) & ")卡类别ID为零(,请与系统管理员联系!": Exit Function
                Set objCard = zlGetCardFromCardType(lng卡类别ID, False, NVL(!结算方式))
                If zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) = False Then Exit Function
                
                objItem.结算类型 = Val(NVL(!类型))
                objItem.结算IDs = objBalanceInfor.结帐ID
                objItem.结帐ID = objBalanceInfor.结帐ID
                objItem.冲销ID = objBalanceInfor.冲销ID
                objItem.结帐时间 = objBalanceInfor.结帐时间
                objItem.结算金额 = RoundEx(intSign * objItem.结算金额, 6)
                objItem.是否允许删除 = Val(NVL(!校对标志)) = 1
                objItem.是否允许退现 = False
                objItem.是否允许编辑 = False
                objItem.预交ID = Val(NVL(!预交ID))
                objItem.是否保存 = bln是否保存
                
                blnAdd = True
                Select Case bytCurType
                Case 0, 1, 2, 4, 6 '0-结帐;1-结帐作废;2-异常重结;3-异常重退,4-正常单据查看;5-作废单据查看;6-异常结算作废
                     
                    If objItem.结算金额 < 0 And objBalanceInfor.冲销ID = 0 Then '结帐信息
                        
                        objItem.是否预交 = True: objItem.是否退款 = True
                        '0-普通业务;1-分交易退款,2-调用一次交易接口退款;3-转帐方式退款
                        Select Case Val(NVL(!附加标志))
                        Case 2 '调用一次交易接口退款
                            If InStr(strMulitCardTypeIds & ",", "," & lng卡类别ID & ",") = 0 Then strMulitCardTypeIds = strMulitCardTypeIds & "," & lng卡类别ID
                            blnFind = False
                            For i = 1 To objBalanceItems_out.Count
                                 If objItem.卡类别ID = objBalanceItems_out(i).卡类别ID Then
                                     If objBalanceItems_out(i).objTag Is Nothing Then
                                         Set objBalanceItems_out(i).objTag = New clsBalanceItems
                                     End If
                                     Set objItems = objBalanceItems_out(i).objTag
                                     objItems.AddItem objItem
                                     objItems.结算金额 = RoundEx(objItems.结算金额 + objItem.结算金额, 6)
                                     objBalanceItems_out(i).结算金额 = RoundEx(objBalanceItems_out(i).结算金额 + objItem.结算金额, 6)
                                     objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
                                     blnFind = True
                                     Exit For
                                 End If
                            Next
                            If Not blnFind Then
                                 Set objItemTemp = zlCopyNewItemFromBalanceItem(objItem)
                                 Set objItemTemp.objTag = New clsBalanceItems
                                 Set objItems = objItemTemp.objTag
                                 objItemTemp.是否退款分交易 = False
                                 objItems.结算金额 = RoundEx(objItems.结算金额 + objItem.结算金额, 6)
                                 objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
                                 objItems.AddItem objItem
                                 objBalanceItems_out.AddItem objItemTemp
                            End If
                            blnAdd = False
                        Case 1   '1-分交易退款
                            strTemp = lng卡类别ID & "|" & objItem.预交ID
                            If InStr(strSingleCardTypeIds & ",", "," & strTemp & ",") = 0 Then strSingleCardTypeIds = strSingleCardTypeIds & "," & strTemp
                            If Not (bytCurType = 4 Or bytCurType = 5) Then  '不等于查看，需要处理转账或分交易退款
                                If zlGetDelBalanceItemsFromRecord(objItem, rsThirdDel, objItemsTemp, True) = False Then Exit Function
                                objItem.是否退款分交易 = True
                                If objItem.预交ID <> objItemsTemp(1).预交ID Then objItem.预交ID = objItemsTemp(1).预交ID
                                Set objItem.objTag = objItemsTemp
                            End If
                        Case Else   '0-普通业务;1-分交易退款,2-调用一次交易接口退款;3-转帐方式退款
                            If InStr(strMulitCardTypeIds & ",", "," & lng卡类别ID & ",") = 0 Then strMulitCardTypeIds = strMulitCardTypeIds & "," & lng卡类别ID
                            If Not (bytCurType = 4 Or bytCurType = 5) Then  '不等于查看，需要处理转账或分交易退款
                                If zlGetDelBalanceItemsFromRecord(objItem, rsThirdDel, objItemsTemp) = False Then Exit Function
                                Set objItem.objTag = objItemsTemp
                            End If
                        End Select
                    ElseIf objItem.结算金额 > 0 And bytCurType = 1 Then '结帐作废时，需要先检查是否缺省
                            
                        If InStr(strMulitCardTypeIds & ",", "," & lng卡类别ID & ",") = 0 Then strMulitCardTypeIds = strMulitCardTypeIds & "," & lng卡类别ID
                        
                        strCardTypes = !类型 & "_" & Val(NVL(!卡类别ID)) & "_" & Val(NVL(!关联交易ID))
                        Set objItemsTemp = New clsBalanceItems
                        i = 0
                        objItem.是否退款 = True
                        Do While Not .EOF
                            If strCardTypes <> !类型 & "_" & Val(NVL(!卡类别ID)) & "_" & Val(NVL(!关联交易ID)) Then .MovePrevious: Exit Do
                            If i <> 0 Then
                                Set objItem = New clsBalanceItem
                                If zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) = False Then Exit Function
                                objItem.结算类型 = Val(NVL(!类型))
                                objItem.结算IDs = objBalanceInfor.结帐ID
                                objItem.结帐ID = objBalanceInfor.结帐ID
                                objItem.冲销ID = objBalanceInfor.冲销ID
                                objItem.结帐时间 = objBalanceInfor.结帐时间
                                objItem.是否允许删除 = False
                                objItem.是否允许退现 = False
                                objItem.是否允许编辑 = False
                                objItem.是否退款 = True
                                objItem.是否保存 = bln是否保存
                                objBalanceInfor.当前结帐 = RoundEx(objBalanceInfor.当前结帐 + objItem.结算金额, 6)
                            End If
                            i = i + 1
                            objItemsTemp.AddItem objItem
                            objItemsTemp.结算金额 = objItemsTemp.结算金额 + objItem.结算金额
                            .MoveNext
                        Loop
                        If .EOF Then .MovePrevious
                        
                        For i = 1 To objItemsTemp.Count
                            '金额写反数,避免传入接口为负数
                            objItemsTemp(i).结算金额 = RoundEx(-1 * objItemsTemp(i).结算金额, 6)
                        Next
                        '启用参数"按结帐金额产生新的预交款"时,不调用退现接口
                        If Not (bytCurType = 1 And gTy_System_Para.TY_Balance.bln结帐作废产生新预交 And gTy_System_Para.TY_Balance.str作废预交结算方式 <> "") Then
                            blnReturnCash = objThirdSwap.zlThirdReturnCashCheck(objCard, objItemsTemp, blnDelCash, strDefaultBalance)
                        End If
                        If Not blnReturnCash Then
                            For i = 1 To objItemsTemp.Count
                                objItemsTemp(i).是否允许退现 = False
                                objItemsTemp(i).是否强制退现 = blnDelCash
                                objItemsTemp(i).是否允许删除 = objItemsTemp(i).是否强制退现
                                objItemsTemp(i).结算金额 = RoundEx(-1 * objItemsTemp(i).结算金额, 6)
                                objItemsTemp(i).是否退款 = True
                                objItemsTemp(i).结帐ID = objBalanceInfor.结帐ID
                                objItemsTemp(i).冲销ID = objBalanceInfor.冲销ID
                                objItemsTemp(i).结帐时间 = objBalanceInfor.结帐时间
                                objItemsTemp(i).是否保存 = bln是否保存
                                objBalanceItems_out.AddItem objItemsTemp(i)
                                objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + objItemsTemp(i).结算金额, 6)
                                objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItemsTemp(i).结算金额, 6)
                            Next
                            blnAdd = False
                        Else
                            For i = 1 To objItemsTemp.Count
                                '金额写反数,避免传入接口为负数
                                objItemsTemp(i).结算金额 = RoundEx(-1 * objItemsTemp(i).结算金额, 6)
                            Next
                        
                            blnAdd = False
                            If blnDelCash = False Then  '是否缺省退现
                                For i = 1 To objItemsTemp.Count
                                    objItemsTemp(i).是否允许退现 = True
                                    objItemsTemp(i).是否允许删除 = True
                                    objItemsTemp(i).是否强制退现 = True
                                    objItemsTemp(i).结帐ID = objBalanceInfor.结帐ID
                                    objItemsTemp(i).冲销ID = objBalanceInfor.冲销ID
                                    objItemsTemp(i).结帐时间 = objBalanceInfor.结帐时间
                                    objItemsTemp(i).是否保存 = bln是否保存
                                    
                                    objBalanceItems_out.AddItem objItemsTemp(i)
                                    objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + objItemsTemp(i).结算金额, 6)
                                    objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItemsTemp(i).结算金额, 6)
                                Next
                           ElseIf strDefaultBalance <> "" Then
                                Set objItemTemp = New clsBalanceItem
                                With objItemTemp
                                    Set .objCard = zlGetCardFromBalanceName(strDefaultBalance)
                                    .结算方式 = strDefaultBalance
                                    .结算金额 = RoundEx(intSign * objItemsTemp.结算金额, 6)
                                    .是否退款 = True
                                    .是否允许编辑 = False
                                    .是否允许删除 = True
                                    .结算性质 = .objCard.结算性质
                                    .结算IDs = objBalanceInfor.结帐ID
                                    .结帐ID = objBalanceInfor.结帐ID
                                    .冲销ID = objBalanceInfor.冲销ID
                                    .结帐时间 = objBalanceInfor.结帐时间
                                    .是否保存 = bln是否保存
                                End With
                                objPtItems.AddItem objItemTemp
                                objPtItems.结算金额 = RoundEx(objPtItems.结算金额 + objItem.结算金额, 6)
                            End If
                         
                        End If
                    End If
                End Select
                
                If blnAdd Then
                    objBalanceItems_out.AddItem objItem
                    objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + objItem.结算金额, 6)
                    objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
                End If
                
            Case 4 '旧一卡通
                
                Set objCard = zlGetCardFromCardType(lng卡类别ID, False, NVL(!结算方式))
                If zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) = False Then Exit Function
                
                objItem.结算类型 = Val(NVL(!类型))
                objItem.结算IDs = objBalanceInfor.结帐ID
                objItem.结帐ID = objBalanceInfor.结帐ID
                objItem.冲销ID = objBalanceInfor.冲销ID
                objItem.结帐时间 = objBalanceInfor.结帐时间
                objItem.是否保存 = bln是否保存
                objItem.结算金额 = RoundEx(intSign * objItem.结算金额, 6)
                If objItem.结算金额 < 0 And objBalanceInfor.冲销ID = 0 Then  '三方退款
                    '获取明细集
                    objItem.是否预交 = True: objItem.是否退款 = True
                End If
                
                objItem.是否保存 = bln是否保存
                objBalanceItems_out.AddItem objItem
                objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + objItem.结算金额, 6)
                objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
            Case 5 '消费卡
                lng卡类别ID = Val(NVL(!结算卡序号))
                Set objCard = zlGetCardFromCardType(lng卡类别ID, True, NVL(!结算方式))
                If zlGetBalanceItemFromRecord(objCard, rsBalanceRecord, objItem, strErrMsg_out) = False Then Exit Function
                objItem.卡类别ID = lng卡类别ID
                objItem.结算类型 = Val(NVL(!类型))
                objItem.结帐ID = objBalanceInfor.结帐ID
                objItem.结算IDs = objBalanceInfor.结帐ID
                objItem.冲销ID = objBalanceInfor.冲销ID
                Select Case bytCurType
                Case 0, 1, 2, 4, 6 '0-结帐;1-结帐作废;2-异常重结;3-异常重退,4-正常单据查看;5-作废单据查看;6-异常结算作废
                    If objItem.结算金额 < 0 And objBalanceInfor.冲销ID = 0 Then objItem.是否预交 = True
                End Select
                objItem.消费卡 = True
                objItem.结帐时间 = objBalanceInfor.结帐时间
                objItem.是否保存 = bln是否保存
                objItem.结算金额 = RoundEx(intSign * objItem.结算金额, 6)
                objBalanceItems_out.AddItem objItem
                objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + objItem.结算金额, 6)
                objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
            Case Else
                '可能是误差费，不处理
            End Select
            rsBalanceRecord.MoveNext
        Loop
    End With
    
    '结算信息中可能不存在三方卡但退款信息中，默认应该读取出来
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    Set cllDelSwap = New Collection

    If Not (bytCurType = 4 Or bytCurType = 5 Or bytCurType = 6) Then '不为查看或异常作废时，不用加载这部分无用数据

        rsThirdDel.Filter = ""
        Set rsThirdDelClone = zlDatabase.CopyNewRec(rsThirdDel)

        With rsThirdDel
            .Filter = 0
            Do While Not .EOF

                lng卡类别ID = Val(NVL(!卡类别ID))
                bln转帐 = Val(NVL(rsThirdDel!是否转帐)) = 1
                strTemp = lng卡类别ID & "|" & Val(NVL(!预交ID))

                If Not (InStr(1, strMulitCardTypeIds & ",", "," & lng卡类别ID & ",") > 0 Or InStr(strSingleCardTypeIds & ",", "," & strTemp & ",") > 0) Then

                    Set objCard = zlGetCardFromCardType(lng卡类别ID, False, "")
                    Set objItem = New clsBalanceItem
                    With objItem
                        .关联交易ID = 0
                        .结算金额 = 0
                        .结算方式 = objCard.结算方式
                        .结算类型 = 3

                        .结算IDs = objBalanceInfor.结帐ID
                        .结帐ID = objBalanceInfor.结帐ID
                        .冲销ID = objBalanceInfor.冲销ID
                        .结算性质 = objCard.结算性质
                        .结帐时间 = objBalanceInfor.结帐时间
                        .卡类别ID = lng卡类别ID
                        .是否保存 = bln是否保存
                        .是否退款 = True
                        .是否结算 = Val(NVL(rsThirdDel!是否未退)) = 0
                        .是否预交 = True
                        .是否密文 = objCard.卡号密文规则 <> ""
                        .校对标志 = IIf(Val(NVL(rsThirdDel!是否未退)) = 1, 1, 2)
                        .是否转帐 = bln转帐
                        .预交ID = Val(NVL(rsThirdDel!预交ID))
                        .是否允许删除 = Val(NVL(rsThirdDel!是否未退)) = 1

                        Set .objCard = objCard
                    End With
                    
                    blnNoBalanceData = True
                    If Not (bln转帐 Or InStr(1, strSingleCardTypeIds & "|", "," & lng卡类别ID & "|") > 0) Then
                        blnFind = False
                        For i = 1 To cllDelSwap.Count
                           If cllDelSwap(i)(0) = lng卡类别ID Then
                                blnFind = True
                                blnSingleDel = cllDelSwap(i)(1) = 0: Exit For
                           End If
                        Next

                        If Not blnFind And (bytCurType = 2 Or bytCurType = 3) Then
                            blnSingleDel = ThirdSwapIsSwapNOCall(objBalanceInfor.结帐ID, lng卡类别ID, blnNoBalanceData)
                            If blnNoBalanceData Then
                                '需要重新确定是否分交易调用
                                blnSingleDel = objThirdSwap.zlThirdSwapIsSwapNOCall(lng卡类别ID, False, strErrMsg, strExpend)
                            End If
                            cllDelSwap.Add Array(lng卡类别ID, IIf(blnSingleDel, 0, 1))
                        End If
                        objItem.是否退款分交易 = blnSingleDel
                    ElseIf InStr(1, strSingleCardTypeIds & "|", "," & lng卡类别ID & "|") > 0 Then
                        objItem.是否退款分交易 = True
                    End If

                    If objItem.是否退款分交易 Then
                        objItem.关联交易ID = Val(NVL(rsThirdDel!关联交易ID))
                        objItem.交易流水号 = NVL(rsThirdDel!交易流水号)
                        objItem.交易说明 = NVL(rsThirdDel!交易说明)
                    End If

                    If zlGetDelBalanceItemsFromRecord(objItem, rsThirdDelClone, objItemsTemp, objItem.是否退款分交易) = False Then Exit Function

                    Set objItem.objTag = objItemsTemp
                    objItem.结算金额 = objItemsTemp.结算金额

                    '没有病人预交记录的都调用
                    If (bytCurType = 2 Or bytCurType = 3) And blnNoBalanceData Then
                        '启用参数"按结帐金额产生新的预交款"时,不调用退现接口
                        If Not (bytCurType = 3 And gTy_System_Para.TY_Balance.bln结帐作废产生新预交 And gTy_System_Para.TY_Balance.str作废预交结算方式 <> "") Then
                            blnReturnCash = objThirdSwap.zlThirdReturnCashCheck(objCard, objItemsTemp, blnDelCash, strDefaultBalance)
                        End If
                        If Not blnReturnCash Then
                            objItem.是否允许退现 = False
                            objItem.是否强制退现 = blnDelCash
                            objItem.是否允许删除 = objItem.是否强制退现
                            blnAdd = True
                        Else
                            objItem.是否允许退现 = True: objItem.是否强制退现 = True
                            objItem.是否允许删除 = True
                            If blnDelCash = False Then  '是否缺省退现
                                objItem.是否允许编辑 = False
                                objItem.是否允许删除 = True
                                blnAdd = True
                            ElseIf strDefaultBalance <> "" Then
                                Set objItemTemp = New clsBalanceItem
                                With objItemTemp
                                    Set .objCard = zlGetCardFromBalanceName(strDefaultBalance)
                                    .结算方式 = strDefaultBalance
                                    .结算金额 = RoundEx(objItem.结算金额, 6)
                                    .是否退款 = True
                                    .是否允许编辑 = False
                                    .是否允许删除 = True
                                    .结算性质 = .objCard.结算性质
                                    .结算IDs = objBalanceInfor.结帐ID
                                    .结帐ID = objBalanceInfor.结帐ID
                                    .冲销ID = objBalanceInfor.冲销ID
                                    .结帐时间 = objBalanceInfor.结帐时间
                                    .是否允许退现 = True
                                    .是否强制退现 = True
                                End With
                                objPtItems.AddItem objItemTemp
                                objPtItems.结算金额 = RoundEx(objPtItems.结算金额 + objItemTemp.结算金额, 6)
                                blnAdd = False
                            End If
                        End If
                    Else
                        blnAdd = True
                    End If

                    If blnAdd Then
                        objBalanceItems_out.AddItem objItem
                    End If

                    If objItem.是否退款分交易 Then
                        strSingleCardTypeIds = strSingleCardTypeIds & "," & strTemp
                    Else
                        strMulitCardTypeIds = strMulitCardTypeIds & "," & lng卡类别ID
                    End If
                End If
                .MoveNext
            Loop
        End With
    End If
    
    '加上普通的结算方式
    For Each objItem In objPtItems
        blnAdd = True
        For Each objItemTemp In objBalanceItems_out
            If objItemTemp.结算方式 = objItem.结算方式 And objItemTemp.结算类型 = 0 Then
                objItemTemp.结算金额 = RoundEx(objItemTemp.结算金额 + objItem.结算金额, 6)
                objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
                blnAdd = False
                Exit For
            End If
        Next
        If blnAdd Then
            objBalanceItems_out.AddItem objItem
            objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
        End If
    Next
    objBalanceInfor.未付合计 = RoundEx(objBalanceInfor.当前结帐 - objBalanceInfor.已付合计, 6)
    objBalanceInfor.是否保存结帐单 = bln是否保存
    
    zlGetBalanceItemsFromRecord = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ThirdSwapIsSwapNOCall(ByVal lng结帐ID As Long, ByVal lng卡类别ID As Long, ByRef blnNoData As Boolean) As Boolean
    '判断是否分单据交易
    '入参:
    '出参:
    '   blnNoData-是否无结算数据
    '说明:
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    blnNoData = False
    strSQL = _
        " Select 附加标志 From 病人预交记录" & _
        " Where 记录性质 = 2 And 结帐id = [1] And 卡类别id = [2] And 冲预交 < 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否分单据交易", lng结帐ID, lng卡类别ID)
    If rsTemp.EOF Then blnNoData = True: Exit Function
    ThirdSwapIsSwapNOCall = Val(NVL(rsTemp!附加标志)) = 1 '分交易退款
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
    '功能:根据结帐记录数据返回指定的结算信息集
    '入参:objCard-卡对象
    '     int类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    '     lng关联交易Id-关联交易ID
    '出参:objBalanceItem_Out-返回结算信息集
    '     strErrMsg_Out-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-30 10:31:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Dim dblMoney As Double
    
    Set objBalanceItem_Out = New clsBalanceItem
    
    On Error GoTo errHandle
    
    strTemp = "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!"
    If rsBalanceRecord Is Nothing Then strErrMsg_out = strTemp: Exit Function
    If rsBalanceRecord.State <> 1 Then strErrMsg_out = strTemp: Exit Function
    If rsBalanceRecord.EOF Then Exit Function
    With rsBalanceRecord
        With objBalanceItem_Out
            Set .objCard = objCard
            .结算方式 = NVL(rsBalanceRecord!结算方式)
            .结算金额 = Val(NVL(rsBalanceRecord!冲预交))
            .关联交易ID = Val(NVL(rsBalanceRecord!关联交易ID))
            .交易流水号 = NVL(rsBalanceRecord!交易流水号)
            .交易说明 = NVL(rsBalanceRecord!交易说明)
            .结算号码 = NVL(rsBalanceRecord!结算号码)
            .结算性质 = Val(NVL(rsBalanceRecord!性质))
            .结算摘要 = NVL(rsBalanceRecord!摘要)
            .卡号 = NVL(rsBalanceRecord!卡号)
            .卡类别ID = Val(NVL(rsBalanceRecord!卡类别ID))
            .消费卡ID = Val(NVL(rsBalanceRecord!消费卡ID))
            .消费卡 = Val(NVL(rsBalanceRecord!消费卡ID)) <> 0
            .是否密文 = Val(NVL(rsBalanceRecord!是否密文)) = 1
            .原始金额 = Val(NVL(rsBalanceRecord!冲预交))
            .未退金额 = Val(NVL(rsBalanceRecord!冲预交))
            .是否退款 = Val(NVL(rsBalanceRecord!冲预交)) < 0
            .是否允许编辑 = False
            .是否允许删除 = False
            .校对标志 = Val(NVL(rsBalanceRecord!校对标志))
            .是否结算 = Val(NVL(rsBalanceRecord!校对标志)) = 2 Or Val(NVL(rsBalanceRecord!校对标志)) = 0
            '.结帐时间 =  Format(rsBalanceRecord!收款时间, "yyyy-mm-dd HH:MM:SS")
            .限制类别 = "" '  Nvl(rsBalanceRecord!限制类别)
            .密码 = ""
            .帐户余额 = 0
            .结算类型 = Val(NVL(rsBalanceRecord!类型))
            .结算IDs = 0 'Val(nvl(rsBalanceRecord!结帐ID))
        End With
    End With
    zlGetBalanceItemFromRecord = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlInsureCheck(ByVal str保险结算 As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前的医保是否需要较对
    '入参:str保险结算-保险结算
    '       strAdvance-医保返回的结算
    '出参:
    '返回:需要较对,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo errHandle
    If Not (strAdvance <> "" And str保险结算 <> strAdvance) Then Exit Function
    '正式结算前后,结算方式和结算金额未发生变化时不校对
    blnMedicareCheck = True
    varData = Split(str保险结算, "||"): varData1 = Split(strAdvance, "||")

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

Public Function zlCheckInsureCancelIsValied(ByVal lng结帐ID As Long, ByVal str作废结算方式s As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的结算结算方式检查是否原样退
    '入参:lng结帐ID-结帐ID
    '     str结算方式s-本次作废的结算信息,多个用逗号分离
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-17 13:45:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str结算信息 As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select  结算方式 From 病人预交记录 A,结算方式 B  " & vbCrLf & _
    "   Where a.结帐ID=[1] and a.结算方式=B.名称 and b.性质 in (3,4)  and mod( A.记录性质,10)<>1  " & vbCrLf & _
    "         And nvl(a.卡类别ID,0)=0 And 结算方式 not in (Select Column_value From table(f_str2List([2])))" & vbCrLf
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取医保是否包含原始结算信息", lng结帐ID, str作废结算方式s)
    If rsTemp.EOF Then zlCheckInsureCancelIsValied = True: Exit Function
    
    str结算信息 = ""
    With rsTemp
        Do While Not .EOF
            str结算信息 = str结算信息 & vbCrLf & NVL(!结算方式)
            .MoveNext
        Loop
    End With
    MsgBox "该医保不支持原样退回处理,不允许作废，不支持作废的结算信息如下:" & str结算信息, vbInformation + vbOKOnly, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If


End Function
Public Function zlGetBalanceItemFromCardObject(ByVal objCurCard As Card, ByVal dblMoney As Double, ByRef objItem_Out As clsBalanceItem, _
    Optional str结算摘要 As String, Optional str结算号码 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡对象，获取新的结算信息对象
    '入参:objCurCard-当前卡对象
    '     dblMoney-当前结算金额
    '出参:objItem_Out-当前操作对象
    '     objCurCard-返回当前的支付卡类别
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-23 19:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, int类型    As Integer
    On Error GoTo errHandle
    
    If objCurCard Is Nothing Then Exit Function
    
    int类型 = IIf(objCurCard.接口序号 > 0, IIf(objCurCard.消费卡, 5, 3), 0)
    If objCurCard.结算性质 = 7 Then int类型 = 4
    
    Set objItem_Out = New clsBalanceItem
    With objItem_Out
        Set .objCard = objCurCard: Set .objTag = Nothing
        .关联交易ID = 0
        .结算IDs = ""
        .结算方式 = objCurCard.结算方式
        .结算号码 = str结算号码
        .结算金额 = dblMoney
        .结算类型 = int类型  '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        .结算性质 = objCurCard.结算性质
        .结算摘要 = str结算摘要
        .卡号 = ""
        .卡类别ID = IIf(objCurCard.接口序号 > 0, objCurCard.接口序号, 0)
        .是否密文 = objCurCard.卡号密文规则 <> ""
        .是否允许编辑 = False
        .是否允许删除 = False
        .是否允许退现 = False
        .是否转帐 = False
        .限制类别 = ""
        .消费卡 = objCurCard.消费卡
        .校对标志 = 1
    End With
    zlGetBalanceItemFromCardObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceItemsFromVsBalanceGrid(ByVal vsBalance As VSFlexGrid, ByVal objCurItem As clsBalanceItem, ByRef objItems_Out As clsBalanceItems, _
    Optional ByVal blnNewItem As Boolean = False, Optional blnSign As Boolean, Optional str结算方式s_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算网格，获取指定数据集
    '入参:vsBalance-结算列表
    '     objItem-当前结算信息对象
    '     blnNewItem-是否返回的值是全新的值
    '     blnSign-金额是否取相反数,true-取相返数,否则原始值
    '出参:objItems_Out-返回的关联集
    '      str结算方式s_out:返回结算信息:格式:结算方式:结算金额|结算方式:结算金额;
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-13 16:56:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As clsBalanceItem, objNewItem As clsBalanceItem
    Dim intSign As Integer, str结算方式s As String
    On Error GoTo errHandle
    
    intSign = IIf(blnSign, -1, 1)
    str结算方式s_out = ""
    Set objItems_Out = New clsBalanceItems
    If objCurItem.关联交易ID = 0 Then  '无关联交易ID,则直接返回s
        If blnNewItem Then
            Set objNewItem = zlCopyNewItemFromBalanceItem(objCurItem)
        Else
            Set objNewItem = objCurItem
        End If
        objNewItem.结算金额 = RoundEx(intSign * objNewItem.结算金额, 6)
        str结算方式s_out = objNewItem.结算方式 & ":" & Format(objNewItem.结算金额, "0.00")
        objItems_Out.AddItem objNewItem
        objItems_Out.结算金额 = objNewItem.结算金额
        zlGetBalanceItemsFromVsBalanceGrid = True
        
        Exit Function
    End If
    With vsBalance
        For i = 1 To .Rows - 1
             If zlGetBalanceItemFromBalanceGrid(vsBalance, i, objItem) Then
                If objCurItem.关联交易ID = objItem.关联交易ID And objCurItem.卡类别ID = objItem.卡类别ID And objCurItem.预交ID = objItem.预交ID Then
                    If blnNewItem Then
                        Set objNewItem = zlCopyNewItemFromBalanceItem(objItem)
                    Else
                        Set objNewItem = objItem
                    End If
                    objNewItem.结算金额 = RoundEx(intSign * objNewItem.结算金额, 6)
                    objItems_Out.AddItem objNewItem
                    objItems_Out.是否转帐 = objNewItem.是否转帐
                    objItems_Out.收费类型 = IIf(objNewItem.是否预交, 1, 0)
                    objItems_Out.收费类型 = objNewItem.结算类型
                    objItems_Out.结算金额 = RoundEx(objItems_Out.结算金额 + objNewItem.结算金额, 6)
                    str结算方式s_out = str结算方式s_out & "|" & objNewItem.结算方式 & ":" & Format(objNewItem.结算金额, "0.00")
                End If
             End If
        Next
        If str结算方式s_out <> "" Then str结算方式s_out = Mid(str结算方式s_out, 2)
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
    '功能:获取结帐作废的普通结算方式
    '入参:bytFun-0-普通;1-医保;2-消费卡
    '     vsBalance-结算列表
    '出参:
    '    bytfunc=0:strBalances的格式:结算方式|结算金额|结算号码||...
    '    bytfunc=1:strBalances的格式:结算方式|结算金额||...
    '    bytfunc=2:strBalances的格式:卡类别ID|卡号|消费卡ID|消费金额||.
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-22 16:20:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPTBalance As String, i As Long, dblMoney As Double
    Dim strYbBalance As String, strBalance As String, varData As Variant
    Dim strXFBalance As String
    Dim objItem As clsBalanceItem
    
    On Error GoTo errHandle
    With vsBalance
        '收集退款方式及金额
        strPTBalance = "": strYbBalance = "": strXFBalance = ""
        For i = 1 To .Rows - 1
            dblMoney = -1 * RoundEx(Val(.TextMatrix(i, .ColIndex("结算金额"))), 6)
            strBalance = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            
            If strBalance <> "" And Val(.TextMatrix(i, .ColIndex("结算状态"))) = 0 Then
                '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                Select Case Val(.TextMatrix(i, .ColIndex("类型")))
                Case 0 '普通结算
                    '结算方式|结算金额|结算号码|结算摘要||..
                    strPTBalance = strPTBalance & "||" & strBalance
                    strPTBalance = strPTBalance & "|" & dblMoney
                    strPTBalance = strPTBalance & "|" & IIf(.TextMatrix(i, .ColIndex("结算号码")) = "", " ", .TextMatrix(i, .ColIndex("结算号码")))
                    strPTBalance = strPTBalance & "|" & IIf(.TextMatrix(i, .ColIndex("备注")) = "", " ", .TextMatrix(i, .ColIndex("备注")))
                Case 1 '预交款
                Case 2 '医保
                        '结算方式|结算金额||...
                        strYbBalance = strYbBalance & "||" & .TextMatrix(i, .ColIndex("结算方式")) & "|" & dblMoney
                Case 3 '一卡通
                Case 4 '一卡通(老版本)
                Case 5 '消费卡
                
                    If zlGetBalanceItemFromBalanceGrid(vsBalance, i, objItem) = False Then Exit Function
                    
                    '卡类别ID|卡号|消费卡ID|消费金额||.
                    strXFBalance = strXFBalance & "||" & objItem.卡类别ID  ' Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                    strXFBalance = strXFBalance & "|" & IIf(objItem.卡号 = "", " ", objItem.卡号) ' Trim(.Cell(flexcpData, i, .ColIndex("卡号")))
                    strXFBalance = strXFBalance & "|" & objItem.消费卡ID  'Val(.TextMatrix(i, .ColIndex("消费卡ID")))
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
    '功能:根据结算信息,增加结算数据给网格
    '入参:objCard-当前的卡对象
    '     objBalanceItems-当前的结算信息集
    '     lngRow-指定的行(如果为0表示从最后一行加起
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-10 11:38:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    Dim lngTemp As Long
    
    On Error GoTo errHandle
    
    If objCard.消费卡 Then Call zlClearSquareBalance(vsBalance, objCard.接口序号, objBalanceInfor)        '消费卡，需要清除原已经存在的数据
    If lngRow <= 0 Then lngRow = zlGetBalanceNULLRow(vsBalance, lngRow)
    If lngRow < 0 Then vsBalance.Rows = vsBalance.Rows + 1: lngRow = vsBalance.Rows - 1
    
    If objBalanceItems Is Nothing Then Exit Sub
    
    With vsBalance
        If .Rows <= 1 Then .Rows = 2
        If lngRow > .Rows - 1 Then
             If Trim(.TextMatrix(.Rows - 1, .ColIndex("结算方式"))) <> "" Then
                .Rows = .Rows + 1
             End If
             lngRow = .Rows - 1
        End If
        
        For Each objItem In objBalanceItems
              If Trim(.TextMatrix(lngRow, .ColIndex("结算方式"))) <> "" And Val(.TextMatrix(lngRow, .ColIndex("结算状态"))) <> 0 Then
                '已经结算过的数据,应该从当前行插入
                 If lngRow >= .Rows - 1 Then
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                 ElseIf Trim(.TextMatrix(lngRow + 1, .ColIndex("结算方式"))) <> "" And Val(.TextMatrix(lngRow + 1, .ColIndex("结算状态"))) <> 0 Then
                    '下一行是结算了的数据，需要在中间插入行，以便同一次结算放在一起
                    lngTemp = zlGetBalanceNULLRow(vsBalance, lngRow)
                    If lngTemp < 0 Then .Rows = .Rows + 1: lngTemp = .Rows = .Rows - 1
                    .RowPosition(lngTemp) = lngRow + 1
                    lngRow = lngRow + 1
                 Else
                    lngRow = lngRow + 1
                 End If
              End If
              
              objItem.行号 = lngRow
              objItem.QRCode = ""
              '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
              .TextMatrix(lngRow, .ColIndex("类型")) = objItem.结算类型
              .TextMatrix(lngRow, .ColIndex("是否密文")) = IIf(objItem.是否密文, 1, 0)
              .TextMatrix(lngRow, .ColIndex("结算性质")) = objCard.结算性质
              .TextMatrix(lngRow, .ColIndex("编辑状态")) = IIf(objItem.是否允许编辑, 1, 0) & "|" & IIf(objItem.是否允许删除, 1, 0)
              .TextMatrix(lngRow, .ColIndex("结算状态")) = IIf(objItem.校对标志 = 2, 1, 0) '是否已结算:1-已结算;0-未结算
              .TextMatrix(lngRow, .ColIndex("卡类别ID")) = objItem.卡类别ID
              .TextMatrix(lngRow, .ColIndex("消费卡ID")) = objItem.消费卡ID
              .TextMatrix(lngRow, .ColIndex("结算方式")) = objItem.结算方式
              .TextMatrix(lngRow, .ColIndex("卡号")) = objCard.zlCardNOEncrypt(objItem.卡号)
              .TextMatrix(lngRow, .ColIndex("结算金额")) = Format(objItem.结算金额, "0.00")
              .TextMatrix(lngRow, .ColIndex("结算号码")) = objItem.结算号码
              .TextMatrix(lngRow, .ColIndex("备注")) = objItem.结算摘要
              .TextMatrix(lngRow, .ColIndex("交易流水号")) = objItem.交易流水号
              .TextMatrix(lngRow, .ColIndex("交易说明")) = objItem.交易说明
              
              .TextMatrix(lngRow, .ColIndex("是否退现")) = IIf(objCard.是否退现, 1, 0)
              .TextMatrix(lngRow, .ColIndex("是否全退")) = IIf(objCard.是否全退, 1, 0)
              .TextMatrix(lngRow, .ColIndex("卡类别名称")) = objCard.名称
              
              .Cell(flexcpData, lngRow, .ColIndex("结算金额")) = Format(objItem.结算金额, "0.00")
              .Cell(flexcpData, lngRow, .ColIndex("消费卡ID")) = objItem.密码
              .Cell(flexcpData, lngRow, .ColIndex("卡类别ID")) = objItem.限制类别
              .Cell(flexcpData, lngRow, .ColIndex("卡号")) = objItem.卡号
              .Cell(flexcpData, lngRow, .ColIndex("结算状态")) = IIf(objItem.是否保存, 1, 0)
              
                If objItem.是否结算 Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = g_BalanceRow_Color_Succes
                ElseIf objItem.是否保存 Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = g_BalanceRow_Color_Valied
                Else
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = g_BalanceRow_Color_Normal
                End If
              .RowData(lngRow) = objItem
              If lngRow + 1 > .Rows - 1 Then .Rows = .Rows + 1
              lngRow = lngRow + 1
        Next
        
    End With
    Call zlRecalItemObjectRowNo(vsBalance)    '重新刷行对象的行号
   Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlLoadBalanceItemsToVsGrid(ByVal vsGrid As VSFlexGrid, ByVal objBalanceItems As clsBalanceItems, Optional ByVal bln查看 As Boolean, Optional ByVal lngRow As Long = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将指定结算数据加载到网格
    '入参:lngRow-指定的行:>0时，替换当前行,然后从最后一行中加入行
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-29 18:14:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytPreDraw As RedrawSettings
    Dim objBalanceItem As clsBalanceItem
    Dim byt类型 As gBalanceType, i As Long
    Dim objCard As Card
    
    On Error GoTo errHandle
    
    If objBalanceItems Is Nothing Then Exit Function '
    
    bytPreDraw = vsGrid.Redraw
    
     
    With vsGrid
        .Redraw = flexRDNone
        If lngRow >= 0 And lngRow < .Rows - 1 Then
            i = lngRow
        Else
            If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) <> "" Then .Rows = .Rows + 1
            i = .Rows - 1
        End If
        
        For Each objBalanceItem In objBalanceItems
            
            Set objCard = objBalanceItem.objCard
            If objCard Is Nothing Then
                'zlGetCard(ByVal lngCardTypeId As Long, ByVal bln消费卡 As Boolean,    ByRef objCard As Card)
                If objBalanceItem.卡类别ID <> 0 Then
                    Call gobjSquare.objOneCardComLib.zlGetCard(objBalanceItem.卡类别ID, objBalanceItem.消费卡, objCard)
                Else
                    Set objCard = zlGetCardFromBalanceName(objBalanceItem.结算方式)
                End If
                Set objBalanceItem.objCard = objCard
            End If
            
            If objCard Is Nothing Then
                Set objCard = New Card
                With objCard
                    .结算方式 = objBalanceItem.结算方式
                    .结算性质 = objBalanceItem.结算性质
                    .是否退现 = True
                    .是否全退 = False
                    .名称 = ""
                End With
                Set objBalanceItem.objCard = objCard
            End If
            
            .TextMatrix(i, .ColIndex("类型")) = objBalanceItem.结算类型
            .TextMatrix(i, .ColIndex("卡类别ID")) = objBalanceItem.卡类别ID
            .TextMatrix(i, .ColIndex("消费卡ID")) = objBalanceItem.消费卡ID
            .TextMatrix(i, .ColIndex("结算性质")) = objBalanceItem.结算性质
            .TextMatrix(i, .ColIndex("编辑状态")) = IIf(objBalanceItem.是否允许编辑, "1", "0") & "|" & IIf(objBalanceItem.是否允许删除, "1", "0")      '是否允许编辑|是否允许删除
            .TextMatrix(i, .ColIndex("是否退现")) = IIf(objCard.是否退现, 1, 0)
            .TextMatrix(i, .ColIndex("是否全退")) = IIf(objCard.是否全退, 1, 0)
            .TextMatrix(i, .ColIndex("校对标志")) = objBalanceItem.校对标志
            .TextMatrix(i, .ColIndex("是否密文")) = IIf(objBalanceItem.是否密文, 1, 0)
            .TextMatrix(i, .ColIndex("卡类别名称")) = objCard.名称
            .TextMatrix(i, .ColIndex("结算方式")) = objBalanceItem.结算方式
            .TextMatrix(i, .ColIndex("结算金额")) = IIf(objBalanceItem.结算性质 = 9, Format(objBalanceItem.结算金额, "###0.00#####"), Format(objBalanceItem.结算金额, "0.00"))
            .TextMatrix(i, .ColIndex("结算号码")) = objBalanceItem.结算号码
            .TextMatrix(i, .ColIndex("备注")) = objBalanceItem.结算摘要
            .TextMatrix(i, .ColIndex("交易流水号")) = objBalanceItem.交易流水号
            .TextMatrix(i, .ColIndex("交易说明")) = objBalanceItem.交易说明
            .TextMatrix(i, .ColIndex("原预交id")) = objBalanceItem.预交ID
            .TextMatrix(i, .ColIndex("卡号")) = IIf(objBalanceItem.是否密文, String(Len(objBalanceItem.卡号), "*"), objBalanceItem.卡号)
            .Cell(flexcpData, i, .ColIndex("卡号")) = NVL(objBalanceItem.卡号)
            .RowData(i) = objBalanceItem
            If bln查看 Then
                If objBalanceItem.校对标志 = 1 Then    '未执行成功的
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                ElseIf objBalanceItem.校对标志 = 2 Then  '执行成功且当前处于查看的
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                Else
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vsGrid.ForeColor
                End If
            End If
            
            If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) <> "" Then .Rows = .Rows + 1
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
     ByRef objBalanceInfor As clsBalanceInfo, Optional ByVal lng消费卡ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除消费卡结算
    '编制:刘兴洪
    '日期:2015-01-23 14:54:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, j As Long
    With vsBalance
        j = 1
        Do While j <= .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("类型"))) = 5 _
                And Val(.TextMatrix(j, .ColIndex("卡类别ID"))) = lngCardTypeID _
                And (lng消费卡ID = 0 Or (lng消费卡ID <> 0 And Val(.TextMatrix(j, .ColIndex("消费卡ID"))) = lng消费卡ID)) Then
                dblMoney = Val(.TextMatrix(j, .ColIndex("结算金额")))
                
                objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 - dblMoney, 6)
                objBalanceInfor.未付合计 = RoundEx(objBalanceInfor.未付合计 + dblMoney, 6)
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
    '功能:获取结算方式中为NULL的结算方式的行
    '入参:lngRow-当前行后面的行
    '出参:
    '返回:-1表示不存在;>1 表示获取成功的行
    '编制:刘兴洪
    '日期:2018-04-10 14:18:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If lngRow = 0 Then lngRow = 1
    With vsBalance
        For i = lngRow To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("结算方式"))) = "" Then
                zlGetBalanceNULLRow = i: Exit Function
            End If
        Next
    End With
    zlGetBalanceNULLRow = -1
End Function

Public Sub zlRecalItemObjectRowNo(ByVal vsBalance As VSFlexGrid)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新Item对象的行号(以保存行号是正确的)
    '编制:刘兴洪
    '日期:2018-07-13 14:06:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItemTemp As clsBalanceItem
    On Error GoTo errHandle
    With vsBalance
         For i = 1 To .Rows - 1
             If zlGetBalanceItemFromBalanceGrid(vsBalance, i, objItemTemp) Then
                objItemTemp.行号 = i
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
    '功能:根据指定的卡对象，检查指定的结算列表中是否存在
    '入参:objCard-卡对象
    '出参:
    '返回:成在返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-24 14:54:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln消费卡 As Boolean, i As Long
    On Error GoTo errHandle
    If objCard Is Nothing Then Exit Function
    
    With vsBalance
        For i = 1 To .Rows - 1
            If objCard.接口序号 > 0 Then
                bln消费卡 = Val(.TextMatrix(i, .ColIndex("类型"))) = 5
                If objCard.接口序号 = Val(.TextMatrix(i, .ColIndex("卡类别ID"))) And objCard.消费卡 = bln消费卡 Then
                     zlCheckVsBalanceIsExsitsFromCardObject = True: Exit Function
                End If
                
            Else
                If objCard.结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式"))) Then
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
    '功能:初始化网格控件
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-12-29 15:08:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsDeposit
        .Clear
        .Cols = 26: .Rows = 2
        i = 0
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "单据号": i = i + 1
        .TextMatrix(0, i) = "类别": i = i + 1
        .TextMatrix(0, i) = "票据号": i = i + 1
        .TextMatrix(0, i) = "收款日期": i = i + 1
        .TextMatrix(0, i) = "结算方式": i = i + 1
        .TextMatrix(0, i) = "余额": i = i + 1
        .TextMatrix(0, i) = "冲预交": i = i + 1
        .TextMatrix(0, i) = "金额": i = i + 1
        .TextMatrix(0, i) = "预交ID": i = i + 1
        .TextMatrix(0, i) = "编辑状态": i = i + 1
        .TextMatrix(0, i) = "卡类别ID": i = i + 1
        .TextMatrix(0, i) = "是否消费卡": i = i + 1
        .TextMatrix(0, i) = "卡类别名称": i = i + 1
        .TextMatrix(0, i) = "卡号": i = i + 1
        .TextMatrix(0, i) = "交易流水号": i = i + 1
        .TextMatrix(0, i) = "交易说明": i = i + 1
        .TextMatrix(0, i) = "是否退现": i = i + 1
        .TextMatrix(0, i) = "是否全退": i = i + 1
        .TextMatrix(0, i) = "是否缺省退现": i = i + 1
        .TextMatrix(0, i) = "是否转帐及代扣": i = i + 1
        .TextMatrix(0, i) = "关联交易ID": i = i + 1
        .TextMatrix(0, i) = "结算性质": i = i + 1
        .TextMatrix(0, i) = "原始金额": i = i + 1
        .TextMatrix(0, i) = "结算号码": i = i + 1
        .TextMatrix(0, i) = "摘要": i = i + 1
          
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedCols = 1
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            
            ''ColData(i):列设置属性(1-固定,-1-不能选,0-可选)|列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            Select Case .ColKey(i)
            Case "单据号"
                .ColData(i) = "1|0"
                .FixedAlignment(i) = flexAlignRightCenter
            Case "余额"
                 If bytEditType = g_Ed_门诊结帐 Or bytEditType = g_Ed_住院结帐 _
                    Or bytEditType = g_Ed_重新结帐 Then
                    .ColData(i) = "0|0"
                    .ColHidden(i) = False
                 Else
                      .ColHidden(i) = True: .ColData(i) = "-1|1"
                 End If
            Case "冲预交"
                    .ColData(i) = "1|0"
                    .ColHidden(i) = False
            Case "金额"
                 If bytEditType = g_Ed_门诊结帐 Or bytEditType = g_Ed_住院结帐 Or bytEditType = g_Ed_重新结帐 Then
                     .ColHidden(i) = True: .ColData(i) = "0|1"
                 Else
                      .ColHidden(i) = True: .ColData(i) = "-1|0"
                 End If
            Case "卡类别名称", "卡号", "交易流水号", "交易说明", "结算号码", "摘要"
                 .ColHidden(i) = True: .ColData(i) = "0|0"
            Case Else
                If Not .ColKey(i) Like "*ID" Then
                    .ColData(i) = "0|0"
                End If
            End Select
            
            If InStr(",是否消费卡,是否退现,是否全退,是否缺省退现,是否转帐及代扣,编辑状态,结算性质,原始金额,", "," & .ColKey(i) & ",") > 0 Or .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"
            ElseIf .ColKey(i) Like "*额" Or .ColKey(i) Like "*冲预交" Then
                .ColAlignment(i) = flexAlignRightCenter
            End If
        Next
        
        .ExtendLastCol = False
        .ExplorerBar = IIf(blnAllowSort, flexExSort, flexExNone)
        .ColHidden(.ColIndex("类别")) = True
        .ColWidth(.ColIndex("类别")) = 1100
        
        .ColHidden(.ColIndex("票据号")) = True
        .ColWidth(.ColIndex("票据号")) = 1100
        .ColWidth(.ColIndex("收款日期")) = 1200
        .ColWidth(.ColIndex("单据号")) = 1100
        .ColWidth(.ColIndex("结算方式")) = 1400
        .ColWidth(.ColIndex("余额")) = 1100
        .ColWidth(.ColIndex("冲预交")) = 1100
        .ColWidth(.ColIndex("卡类别名称")) = 1800
        .ColWidth(.ColIndex("卡号")) = 1100
        .ColWidth(.ColIndex("交易流水号")) = 1100
        .ColWidth(.ColIndex("交易说明")) = 1600
        .ColWidth(.ColIndex("结算号码")) = 1100
        .ColWidth(.ColIndex("摘要")) = 1600
        .ColHidden(.ColIndex("金额")) = True
        .ColData(.ColIndex("金额")) = "-1|1"
        zl_vsGrid_Para_Restore lngModul, vsDeposit, strFromName, strRegName
        If bytEditType = g_Ed_单据查看 Or bytEditType = g_Ed_结帐作废 Or bytEditType = g_Ed_重新作废 Or bytEditType = g_Ed_取消结帐 Then
            .ColHidden(.ColIndex("余额")) = True: .ColData(.ColIndex("余额")) = "-1|1"
        End If
    End With
    zlInitDepositGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Sub zlInitBalanceGrid(ByVal vsBalance As VSFlexGrid, ByVal lngModul As Long, ByVal strFromName As String, ByVal strRegKey As String, Optional bln查看 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算列表
    '编制:刘兴洪
    '日期:2015-01-23 14:14:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBalance
    
        For i = 1 To .Rows - 1
            .RowData(i) = ""
        Next
        .Clear: .Rows = 2: i = 0: .Cols = 21
        .TextMatrix(0, i) = "卡类别ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "消费卡ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算性质": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "编辑状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "类型": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否退现": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否全退": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "校对标志": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否密文": .ColWidth(i) = 0: i = i + 1
        
        .TextMatrix(0, i) = "结算方式": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "结算金额": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "结算号码": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "卡类别名称": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "卡号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易流水号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易说明": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "备注": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "关联交易ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "原预交ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否转账": .ColWidth(i) = 0: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            Select Case .ColKey(i)
            Case "是否转账", "关联交易ID", "结算性质", "类型", "是否保存", "是否密文", "校对标志", "编辑状态", "是否退现", "是否全退", "结算状态", "是否验证", "原预交ID"
                .ColHidden(i) = True
                .ColData(i) = """-1||1"
            Case "结算金额"
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = """1||0"
            Case .ColIndex("结算方式")
                .ColData(i) = """1||0"
            Case "卡类别名称"
                .ColData(i) = "1||2"
            Case .ColIndex("结算号码")
                .ColData(i) = "1||0"
            Case Else
                .ColData(i) = "1||" & IIf(bln查看, "0", "2")
                
            End Select
            If bln查看 Then .ColData(i) = ""
        Next
        If Not bln查看 Then .Editable = flexEDKbdMouse
    End With
    zl_vsGrid_Para_Restore lngModul, vsBalance, strFromName, strRegKey
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



Public Sub zlAutoRecalFeeBalanceMoney(ByVal vsDetailList As VSFlexGrid, ByVal dbl本次结帐 As Double, ByVal dbl本次款结 As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动计算和分摊结帐金额
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-24 15:59:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Long
    Dim blnAll As Boolean
    
    On Error GoTo errHandle
    
    dblMoney = dbl本次结帐
    blnAll = dbl本次结帐 = dbl本次款结
    With vsDetailList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据")) <> "" Then
                If dblMoney >= Val(.Cell(flexcpData, i, .ColIndex("未结金额"))) And dblMoney <> 0 Or blnAll Then
                    .Cell(flexcpData, i, .ColIndex("结帐金额")) = Val(.Cell(flexcpData, i, .ColIndex("未结金额")))
                    dblMoney = RoundEx(dblMoney - Val(.Cell(flexcpData, i, .ColIndex("结帐金额"))), 6)
                Else
                    If dblMoney = 0 Then
                        .Cell(flexcpData, i, .ColIndex("结帐金额")) = ""
                    Else
                        .Cell(flexcpData, i, .ColIndex("结帐金额")) = dblMoney
                    End If
                    dblMoney = 0
                End If
                .TextMatrix(i, .ColIndex("结帐金额")) = Format(Val(.Cell(flexcpData, i, .ColIndex("结帐金额"))), gstrDec)
            End If
        Next i
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Function zlGetRemainderMoneyToPati(ByVal byt类型 As Byte, ByVal lng病人ID As Long, ByRef objPati As clsPatiInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人余额及费用余额给病人对象
    '入参:objPati-当前的病人
    '     byt类型-1-门诊;2-住院;0-所有费用余额
    '出参:objPati-返回更新的病人余额信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-24 20:16:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If byt类型 = 0 Then
        strSQL = "Select sum(预交余额) As 预交余额,sum(费用余额) As 费用余额 From 病人余额 Where 病人ID= [1] And 性质=1"
    Else
        strSQL = "Select sum(预交余额) as 预交余额,sum(费用余额) as 费用余额 From 病人余额 Where 病人ID= [1] And 性质=1 And 类型= [2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人余额信息", lng病人ID, byt类型)
    
    objPati.预交余额 = Format(Val(NVL(rsTemp!预交余额)), "0.00")
    objPati.费用余额 = Format(Val(NVL(rsTemp!费用余额)), "0.00")
    objPati.预交剩余合计 = Format(Val(NVL(rsTemp!预交余额)) - Val(NVL(rsTemp!费用余额)), "0.00")
    zlGetRemainderMoneyToPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckNoSettlementMoney(ByVal str姓名 As String, _
    ByVal lng病人ID As Long, ByVal strTimes As String, _
    Optional ByVal byt结帐类型 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查门诊留观病人是否存在未结费用金额
    '入参:
    '   lng病人ID 指定病人
    '   strTimes 指定住院次数,多个用英文逗号分隔，为空表示所有住院次数
    '   byt结帐类型 1-门诊结帐;2-住院结帐
    '出参:
    '返回:检查通过返回True,否则返回False
    '说明:针对门诊留观病人，进行住院结帐时，如果存在门诊费用则提示；进行门诊结帐时，如果存在住院费用则必须先结住院费用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strWhere As String
    Dim strTemp As String
    
    On Error GoTo ErrHandler
    If strTimes <> "" Then
        strWhere = " And a.主页id In(Select /*+Cardinality(j,10)*/ Column_Value From Table(f_Num2list([3])) J)"
    End If
    strSQL = "Select a.主页ID,Sum(a.金额) As 未结金额" & _
            " From 病人未结费用 A" & _
            " Where a.病人id=[1] And a.来源途径 = [2]" & strWhere & _
            " Group By a.主页ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人未结费用", lng病人ID, IIf(byt结帐类型 = 1, 2, 1), strTimes)
    If rsTemp.EOF Then zlCheckNoSettlementMoney = True: Exit Function
    
    Do While Not rsTemp.EOF
        strTemp = strTemp & "," & lng病人ID & ":" & Val(NVL(rsTemp!主页ID))
        rsTemp.MoveNext
    Loop
    strTemp = Mid(strTemp, 2)
    
    '检查是否为门诊留观住院
    If zlGetPatiPageInfo(0, strTemp, rsTemp) = False Then Exit Function
    rsTemp.Filter = "病人性质=1"
    If rsTemp.EOF Then zlCheckNoSettlementMoney = True: Exit Function
    
    Do While Not rsTemp.EOF
        strTemp = strTemp & "、" & Val(NVL(rsTemp!主页ID))
        rsTemp.MoveNext
    Loop
    strTemp = Mid(strTemp, 2)
    
    If byt结帐类型 = 2 Then
        MsgBox "病人『" & str姓名 & "』在第" & strTemp & "次住院还存在未结清的门诊费用，注意对其进行门诊结帐！", vbInformation, gstrSysName
    Else
        MsgBox "病人『" & str姓名 & "』在第" & strTemp & "次门诊留观还存在未结清的住院费用，必须先对其进行住院结账！", vbInformation, gstrSysName
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

Public Function zlLoadDetaiFeeToGridFromRecord(ByVal rsFeeList As ADODB.Recordset, ByVal bln门诊 As Boolean, ByVal intInsure As Integer, ByRef vsDetailList As VSFlexGrid, _
    ByVal lngModule As Long, ByVal strFromName As String, strRegKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载费目表数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-05 18:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
  
    On Error GoTo errHandle
    If rsFeeList Is Nothing Then Exit Function
    If rsFeeList.State <> 1 Then Exit Function
    
    If rsFeeList.RecordCount <> 0 Then rsFeeList.MoveFirst
    With vsDetailList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        Do While Not rsFeeList.EOF
            .TextMatrix(.Rows - 1, .ColIndex("日期")) = Format(NVL(rsFeeList!时间), "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("单据")) = NVL(rsFeeList!单据号)
            .TextMatrix(.Rows - 1, .ColIndex("项目")) = NVL(rsFeeList!项目)
            .TextMatrix(.Rows - 1, .ColIndex("未结金额")) = Format(NVL(rsFeeList!未结金额), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("未结金额")) = Val(NVL(rsFeeList!未结金额))
            .TextMatrix(.Rows - 1, .ColIndex("结帐金额")) = Format(NVL(rsFeeList!结帐金额), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("结帐金额")) = Val(NVL(rsFeeList!结帐金额))
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsFeeList!ID, 0)
            .TextMatrix(.Rows - 1, .ColIndex("记录性质")) = Val(NVL(rsFeeList!记录性质))
            .TextMatrix(.Rows - 1, .ColIndex("记录状态")) = IIf(Val(NVL(rsFeeList!记录状态)) = 3, 1, Val(NVL(rsFeeList!记录状态)))
            .TextMatrix(.Rows - 1, .ColIndex("执行状态")) = Val(NVL(rsFeeList!执行状态))
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = Val(NVL(rsFeeList!序号))
            If bln门诊 Then .Cell(flexcpData, .Rows - 1, .ColIndex("序号")) = Val(NVL(rsFeeList!门诊标志))
            .Rows = .Rows + 1
            rsFeeList.MoveNext
        Loop
        .Cell(flexcpBackColor, 1, .ColIndex("结帐金额"), .Rows - 1, .ColIndex("结帐金额")) = IIf(intInsure <> 0, .Cell(flexcpBackColor, 1, .ColIndex("单据")), &HFFFFC0)
        If .TextMatrix(1, .ColIndex("单据")) <> "" Then .Rows = .Rows - 1
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


Public Function zlLoadDetaiFeeToGridFromBalanceID(ByVal lng结帐ID As Long, ByVal vsDetailList As VSFlexGrid, _
    ByVal bln门诊结帐 As Boolean, ByVal bln作废 As Boolean, ByVal lngModule As Long, ByVal strFromName As String, strRegKey As String, Optional blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID来加载费目表数据
    '入参:lng结帐ID-结帐ID
    '     blnNoMoved-是否历史数据转移
    '     bln作废-是否查看作废记录
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 18:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, lngRow As Long, intSign As Integer
    
    On Error GoTo errHandle
    
    intSign = IIf(bln作废, -1, 1)
    
    strSQL = _
    "   Select Mod(A.记录性质,10) as 记录性质, A.NO,A.序号," & _
    "          Max(Decode(a.是否保密,1,'***',b.名称)) as 项目," & _
    "          Max(A.发生时间) As 发生时间, " & _
    "          Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1)," & gbytDec & ")) as 标准金额,Sum(A.结帐金额) as 结帐金额, " & _
    "          Decode(A.记录状态,2,2,1) As 记录状态,Max(a.门诊标志) As 门诊标志 " & _
    "   From 住院费用记录 A,收费项目目录 B" & _
    "   Where A.结帐ID= [1] And A.收费细目ID=B.ID " & _
    "   Group by Mod(A.记录性质,10),A.NO,A.序号,Decode(A.记录状态,2,2,1) "
   
    strSQL = strSQL & " UNION ALL " & vbCrLf & _
        Replace(strSQL, "住院费用记录", "门诊费用记录")

    If blnNOMoved Then
        strSQL = Replace(Replace(strSQL, "住院费用记录", "H住院费用记录"), "门诊费用记录", "H门诊费用记录")
    End If
    
    strSQL = "" & _
    "   Select Max(发生时间) As 发生时间,NO,序号,项目, sum(标准金额) as 标准金额," & _
    "          sum(结帐金额) as 结帐金额,记录状态,Max(门诊标志) As 门诊标志 " & _
    "   From (" & strSQL & ")" & _
    "   Group by NO,序号,项目,记录状态" & _
    "   Order by NO,序号"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据结帐ID统计费用明细", lng结帐ID)
    
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsDetailList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        
        .TextMatrix(0, .ColIndex("未结金额")) = "标准金额"
        .TextMatrix(0, .ColIndex("结帐金额")) = IIf(intSign = -1, "作废金额", "结帐金额")
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("日期")) = Format(NVL(rsTemp!发生时间), "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("单据")) = NVL(rsTemp!NO)
            .TextMatrix(.Rows - 1, .ColIndex("项目")) = NVL(rsTemp!项目)
            .TextMatrix(.Rows - 1, .ColIndex("未结金额")) = Format(NVL(rsTemp!标准金额), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("未结金额")) = Val(NVL(rsTemp!标准金额))
            .TextMatrix(.Rows - 1, .ColIndex("结帐金额")) = Format(intSign * Val(NVL(rsTemp!结帐金额)), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("结帐金额")) = intSign * Val(NVL(rsTemp!结帐金额))
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = Val(NVL(rsTemp!序号))
            If bln门诊结帐 Then
                .Cell(flexcpData, .Rows - 1, .ColIndex("序号")) = Val(NVL(rsTemp!门诊标志))
            End If
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        .Cell(flexcpBackColor, 1, .ColIndex("结帐金额"), .Rows - 1, .ColIndex("结帐金额")) = .Cell(flexcpBackColor, 1, .ColIndex("日期"), 0.1, .ColIndex("日期"))
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

 
Public Function zlLoadFeiMuFeeListToGridFromRecord(ByVal rsFeeList As ADODB.Recordset, ByVal bln门诊 As Boolean, ByVal intInsure As Integer, ByRef vsFeeList As VSFlexGrid, _
    ByVal lngModule As Long, ByVal strFromName As String, ByVal strRegKey As String, ByRef dblMoney_out As Double) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载费目表数据
    '入参:rsFeeList-费用集
    '     bln门诊-是否门诊结帐
    '     intInsure-险类
    '出参:dblMoney_out-返回未结金额合计
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-05 18:00:19
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
           lngRow = .FindRow(NVL(rsFeeList!费目, "未知"), "1", .ColIndex("费目"), , True)
           If lngRow < 0 Then
                If .TextMatrix(1, .ColIndex("费目")) = "" Then
                    lngRow = 1
                Else
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                End If
           End If
           
           If .TextMatrix(1, .ColIndex("费目")) = "" Then lngRow = 1
          .TextMatrix(lngRow, .ColIndex("费目")) = NVL(rsFeeList!费目, "未知")
          
          .Cell(flexcpData, lngRow, .ColIndex("应收金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("应收金额"))) + Val(NVL(rsFeeList!应收金额))
          .TextMatrix(lngRow, .ColIndex("应收金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("应收金额"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("实收金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("实收金额"))) + Val(NVL(rsFeeList!实收金额))
          .TextMatrix(lngRow, .ColIndex("实收金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("实收金额"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("未结金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("未结金额"))) + Val(NVL(rsFeeList!未结金额))
          .TextMatrix(lngRow, .ColIndex("未结金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("未结金额"))), gstrDec)
            
          dblMoney(0) = RoundEx(dblMoney(0) + Val(NVL(rsFeeList!应收金额)), 5)
          dblMoney(1) = RoundEx(dblMoney(1) + Val(NVL(rsFeeList!实收金额)), 5)
          dblMoney(2) = RoundEx(dblMoney(2) + Val(NVL(rsFeeList!未结金额)), 5)
          
            rsFeeList.MoveNext
        Loop
        
        .ColSort(.ColIndex("费目")) = flexSortUseColSort
        If .TextMatrix(1, .ColIndex("费目")) <> "" Then
          .Rows = .Rows + 1: lngRow = .Rows - 1
          .TextMatrix(lngRow, .ColIndex("费目")) = "合计"
          .Cell(flexcpData, lngRow, .ColIndex("应收金额")) = dblMoney(0)
          .TextMatrix(lngRow, .ColIndex("应收金额")) = Format(dblMoney(0), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("实收金额")) = dblMoney(1)
          .TextMatrix(lngRow, .ColIndex("实收金额")) = Format(dblMoney(1), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("未结金额")) = dblMoney(2)
          .TextMatrix(lngRow, .ColIndex("未结金额")) = Format(dblMoney(2), gstrDec)
         
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
Public Function zlLoadFeiMuFeeListToGridFromBalanceID(ByVal lng结帐ID As Long, ByVal vsFeeList As VSFlexGrid, _
    ByVal lngModule As Long, ByVal bln作废 As Boolean, ByRef dblBalanceMoney_Out As Double, ByVal strFromName As String, strRegKey As String, Optional blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID,按收据费目统计费用后，将数据加载到网格
    '入参:lng结帐ID-结帐ID
    '    vsFeeList-费用列表网格
    '出参:dblBalanceMoney_Out-当前结帐合计
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-25 10:00:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long, strSQL As String
    Dim dblMoney(0 To 1) As Double, intSign As Integer
    
    On Error GoTo errHandle
    
    intSign = IIf(bln作废, -1, 1)
    strSQL = "" & _
    "   Select Mod(A.记录性质,10) as 记录性质, A.NO,序号,A.收据费目, " & _
    "          sum(Round(A.标准单价*A.数次*Nvl(A.付数,1)," & gbytDec & ")) as 标准金额,sum(A.结帐金额) as 结帐金额 " & _
    "   From 住院费用记录 A " & _
    "   Where A.结帐ID= [1]  " & _
    "   Group by Mod(A.记录性质,10),A.NO,A.序号,A.收据费目 "
    
   
    strSQL = strSQL & " UNION ALL " & vbCrLf & _
        Replace(strSQL, "住院费用记录", "门诊费用记录")

    If blnNOMoved Then strSQL = Replace(Replace(strSQL, "住院费用记录", "H住院费用记录"), "门诊费用记录", "H门诊费用记录")
    
    strSQL = "" & _
    "   Select 收据费目, sum(标准金额) as 标准金额,sum(结帐金额) as 结帐金额 " & _
    "   From (" & strSQL & ")" & _
    "   Group by 收据费目" & _
    "   Order by 收据费目"
    
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据结帐ID统计结帐费目明细", lng结帐ID)
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    
     dblMoney(0) = 0: dblMoney(1) = 0
    With vsFeeList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        
         lngRow = 1
        
        .TextMatrix(0, .ColIndex("应收金额")) = "标准金额"
        .ColHidden(.ColIndex("实收金额")) = True
        
        Do While Not rsTemp.EOF
          .TextMatrix(lngRow, .ColIndex("费目")) = NVL(rsTemp!收据费目, "未知")
          .Cell(flexcpData, lngRow, .ColIndex("应收金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("应收金额"))) + Val(NVL(rsTemp!标准金额))
          .TextMatrix(lngRow, .ColIndex("应收金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("应收金额"))), gstrDec)
         
          .Cell(flexcpData, lngRow, .ColIndex("未结金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("结帐金额"))) + Val(NVL(rsTemp!结帐金额))
          .TextMatrix(lngRow, .ColIndex("未结金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("结帐金额"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("结帐金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("结帐金额"))) + RoundEx(intSign * Val(NVL(rsTemp!结帐金额)), 6)
          .TextMatrix(lngRow, .ColIndex("结帐金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("结帐金额"))), gstrDec)
          
          dblMoney(0) = dblMoney(0) + Val(NVL(rsTemp!标准金额))
          dblMoney(1) = dblMoney(1) + RoundEx(intSign * Val(NVL(rsTemp!结帐金额)), 6)
          .Rows = .Rows + 1: lngRow = .Rows - 1
          rsTemp.MoveNext
        Loop
        dblMoney(0) = RoundEx(dblMoney(0), 5)
        dblMoney(1) = RoundEx(dblMoney(1), 5)
        
        If .TextMatrix(1, .ColIndex("费目")) <> "" Then
           lngRow = .Rows - 1
          .TextMatrix(lngRow, .ColIndex("费目")) = "合计"
          .Cell(flexcpData, lngRow, .ColIndex("应收金额")) = dblMoney(0)
          .TextMatrix(lngRow, .ColIndex("应收金额")) = Format(dblMoney(0), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("未结金额")) = dblMoney(1)
          .TextMatrix(lngRow, .ColIndex("未结金额")) = Format(dblMoney(1), gstrDec)
         
          .Cell(flexcpData, lngRow, .ColIndex("结帐金额")) = dblMoney(1)
          .TextMatrix(lngRow, .ColIndex("结帐金额")) = Format(dblMoney(1), gstrDec)
         
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

Public Function zlLoadDepositListFromBalanceID(ByVal lng结帐ID As Long, vsDeposit As VSFlexGrid, ByVal blnNOMoved As Boolean, _
    ByRef dblTotal_Out As Double, ByRef rsDeposit_Out As ADODB.Recordset, ByRef intCountBill_Out As Integer, _
    ByVal lngModul As Long, Optional strFormName As String, Optional strRegKey As String, Optional bln作废 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID获取冲预交信息信息并加载到预交列表中
    '入参:lng结帐ID-指定的结帐ID
    '     blnNoMoved-当前是否移动到后备表中
    '     bln作废-是否查看作废单据
    '出参:rsDeposit_Out-返回预交记录集
    '     dblTotal_Out-冲预交总计
    '     intCountBill_Out-涉及的票据张数
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 15:09:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, dblTotal As Double
    Dim intSign As Integer
    
    On Error GoTo errHandle
    dblTotal_Out = 0
    Set rsTemp = GetBalanceDeposit(lng结帐ID, blnNOMoved)
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    intSign = IIf(bln作废, -1, 1)
    With vsDeposit
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        'ID,单据号,票据号,日期,结算方式, 金额
        i = 1
        Do While Not rsTemp.EOF
            .RowData(i) = ""
            .TextMatrix(i, .ColIndex("ID")) = rsTemp!ID
            .TextMatrix(i, .ColIndex("单据号")) = rsTemp!单据号
            .TextMatrix(i, .ColIndex("类别")) = NVL(rsTemp!预交类别)
            .TextMatrix(i, .ColIndex("票据号")) = "" & rsTemp!票据号
            .TextMatrix(i, .ColIndex("收款日期")) = Format(rsTemp!日期, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("结算方式")) = NVL(rsTemp!结算方式)
            .TextMatrix(i, .ColIndex("冲预交")) = Format(intSign * rsTemp!金额, "0.00")
            .TextMatrix(i, .ColIndex("卡类别ID")) = Val(NVL(rsTemp!卡类别ID))
            .TextMatrix(i, .ColIndex("是否消费卡")) = Val(NVL(rsTemp!是否消费卡))
            .TextMatrix(i, .ColIndex("卡类别名称")) = NVL(rsTemp!卡类别名称)
            .TextMatrix(i, .ColIndex("交易流水号")) = NVL(rsTemp!交易流水号)
            .TextMatrix(i, .ColIndex("结算号码")) = NVL(rsTemp!结算号码)
            .TextMatrix(i, .ColIndex("摘要")) = NVL(rsTemp!摘要)
            .TextMatrix(i, .ColIndex("卡号")) = NVL(rsTemp!卡号)
            .TextMatrix(i, .ColIndex("预交ID")) = Val(NVL(rsTemp!ID))
            .TextMatrix(i, .ColIndex("交易说明")) = NVL(rsTemp!交易说明)
            .TextMatrix(i, .ColIndex("是否退现")) = Val(NVL(rsTemp!是否退现))
            .TextMatrix(i, .ColIndex("是否全退")) = Val(NVL(rsTemp!是否全退))
            .TextMatrix(i, .ColIndex("是否缺省退现")) = Val(NVL(rsTemp!是否缺省退现))
            .TextMatrix(i, .ColIndex("是否转帐及代扣")) = Val(NVL(rsTemp!是否转帐及代扣))
            .TextMatrix(i, .ColIndex("结算性质")) = Val(NVL(rsTemp!结算性质))
            .TextMatrix(i, .ColIndex("原始金额")) = Val(NVL(rsTemp!原始金额))
            
            .Rows = .Rows + 1: i = i + 1
            dblTotal = dblTotal + intSign * Val(NVL(rsTemp!金额))
            rsTemp.MoveNext
        Loop
        .Row = 1: .Col = .Cols - 1
        If i > 1 Then .Rows = .Rows - 1
        
        .ColWidth(.ColIndex("收款日期")) = 1305
        .ColWidth(.ColIndex("单据号")) = 1100
        .ColWidth(.ColIndex("结算方式")) = 1400
        .ColWidth(.ColIndex("余额")) = 1100
        .ColWidth(.ColIndex("冲预交")) = 1100
        
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

Public Function zlGetDefaultHospitalizedDate(ByVal lng病人ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID获取上次中途结帐时间
    '入参:lng病人ID-病人ID
    '返回:返回上次中途结帐的结束日期,无中途结帐时,返回空
    '编制:刘兴洪
    '日期:2015-01-06 15:25:02
    '说明:原问题号是30043
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select to_char( Max(结束日期) + 1,'yyyy-mm-dd') as 结束日期 " & _
    "   From 病人结帐记录 " & _
    "   Where  记录状态=1  And 病人iD=[1] and nvl(中途结帐,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据病人ID获取上次中途结帐时间", lng病人ID)
    If rsTemp.EOF Then Exit Function
    zlGetDefaultHospitalizedDate = NVL(rsTemp!结束日期)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsCheck病历已接收(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病历是否已经接收
    '入参:
    '出参:
    '返回:已接收返回True,否则返回False
    '说明:30036
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    On Error GoTo ErrHandler
    strValue = zlGetPatiPageExtendInfo(lng病人ID, lng主页ID, "病历接收")
    zlIsCheck病历已接收 = Val(strValue) = 1
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlSetBalanceRowDataFromItemsObject(ByVal vsBlance As VSFlexGrid, ByVal objItems As clsBalanceItems)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置三方结算的结算状态
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-16 19:35:42
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
    '功能:根据指定项，设置结算数据
    '入参:objItem-指定行
    '编制:刘兴洪
    '日期:2018-07-13 13:51:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItemTemp As clsBalanceItem
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    If objItem Is Nothing Then Exit Sub
    
    lngRow = objItem.行号
    
    
    If lngRow > vsBlance.Rows - 1 Or lngRow < 1 Then Exit Sub
    
    If Not zlGetBalanceItemFromBalanceGrid(vsBlance, lngRow, objItemTemp) Then Exit Sub
    
    With vsBlance
        If objItemTemp.卡类别ID <> objItem.卡类别ID Then Exit Sub
         
         Set objItemTemp = objItem
        .TextMatrix(lngRow, .ColIndex("编辑状态")) = IIf(objItemTemp.是否允许编辑, 1, 0) & "|" & IIf(objItemTemp.是否允许删除, 1, 0)
        .TextMatrix(lngRow, .ColIndex("结算金额")) = Format(objItemTemp.结算金额, "0.00")
        If objItemTemp.是否结算 Then
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbGrayText
        ElseIf objItem.是否保存 Then
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
    Optional blnAllCancel As Boolean, Optional byt校对标志 As Byte = 0, Optional blnDeleteSQL As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结帐取消操作或作废操作的相关Sql
    '入参:objBalanceInfor-结算对象
    '     blnAllCancel-是否全作废
    '     blnDeleteSQL-是否强制获取退费相关SQL
    '出参:cllPro-返回相关保存的SQL
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-25 21:00:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng冲销ID As Long
    Dim i As Long
    
    On Error GoTo errHandle
    
    If cllPro Is Nothing Then Set cllPro = New Collection
    
    If objBalanceInfor.冲销ID <> 0 Then
        If objBalanceInfor.是否保存结帐单 And Not blnDeleteSQL And Not blnDeleteSQL Then zlGetBalanceCancelSQL = True: Exit Function
    End If
    
    '记录集中存在作废过程，则不允许再次作废
    For i = 1 To cllPro.Count
        If InStr(UCase(cllPro(i)), UCase("Zl_病人结帐记录_Cancel")) > 0 Then zlGetBalanceCancelSQL = True: Exit Function
    Next
    
    lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
    With objBalanceInfor
        .冲销ID = lng冲销ID
        .结帐时间 = zlDatabase.Currentdate
    End With
    
    '先退结算记录及费用
    strSQL = "Zl_病人结帐记录_Cancel("
    '  No_In         病人结帐记录.No%Type,
    strSQL = strSQL & "'" & objBalanceInfor.结帐单据号 & "',"
    '  冲销id_In     病人结帐记录.Id%Type,
    strSQL = strSQL & "" & lng冲销ID & ","
    '  操作员编号_In 病人结帐记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人结帐记录.操作员姓名%Type
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  作废时间_In   病人结帐记录.收费时间%Type := Null
    strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')"
    strSQL = strSQL & ")"
    zlAddArray cllPro, strSQL
    
    If Not blnAllCancel Then zlGetBalanceCancelSQL = True: Exit Function
     
    'Zl_病人结帐作废_Modify
    strSQL = "Zl_病人结帐作废_Modify("
    '  操作类型_In   Number,
    strSQL = strSQL & "" & 0 & ","
    '  病人id_In     病人结帐记录.病人id%Type,
    strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & lng冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL,"
    '  缴款_In       病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "NULL,"
    '  找补_In       病人预交记录.找补%Type := Null,
    strSQL = strSQL & "NULL,"
    '  误差金额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  预交金额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '操作员编号_In    病人预交记录.操作员编号%Type := Null,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '收款时间_In      病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '冲预交病人ids_In Varchar2 := Null,
    ' 多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
    strSQL = strSQL & "NULL,"
    '  完成作废_In Number:=0
    strSQL = strSQL & "1,"
    '    校对标志_In  Number := 0,
    strSQL = strSQL & "" & byt校对标志 & ","
    '    关联交易id_In    病人预交记录.Id%Type := Null,
    strSQL = strSQL & "NULL,"
    '    清除原交易_In Number:=0
    strSQL = strSQL & "0)"
    zlAddArray cllPro, strSQL
    zlGetBalanceCancelSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetThirdMoneyInforRecordFromSwapID(ByVal str关联交易IDs As String, ByRef rsSwapRecord_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID,获取相关的结算金额信息集
    '入参:str关联交易IDs-关联交易ID，多个用逗号分离
    '出参:rsSwapRecord_Out-返回关联交易ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-27 17:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strWhere As String, lng关联交易ID As Long
    
    On Error GoTo errHandle
    
    If InStr(str关联交易IDs, ",") > 0 Then
        strWhere = "And A.关联交易ID In (Select column_value From table(f_num2List([1])) "
        
    Else
       lng关联交易ID = Val(str关联交易IDs)
       strWhere = " And  A.关联交易ID =[2]"
    End If
    
    strSQL = "" & _
    "   Select 关联交易ID,卡类别ID,结算方式,交易流水号,交易说明, " & vbCrLf & _
    "          nvl(金额,0)+decode(mod(记录性质,10),1,0,1)* decode(sign(nvl(冲预交,0)),1,1,0)* nvl(冲预交,0) as 原始金额, " & _
    "          decode(sign(nvl(金额,0)),-1,1,0)*nvl(金额,0)+ decode(sign(nvl(冲预交,0)),-1,1,0)* nvl(冲预交,0) as 已退金额" & _
    "   From 病人预交记录 A " & _
    "   Where 1=1 " & strWhere & _
    "   Union all " & _
    "   Select a.关联交易ID,a.卡类别ID,a.结算方式,a.交易流水号,a.交易说明, " & vbCrLf & _
    "          0 as 原始金额, " & _
    "         -1*nvl(b.金额,0) as 已退金额" & _
    "   From 病人预交记录 A,三方退款信息 B" & _
    "   Where  a.ID=b.记录ID And b.是否转帐 =1  " & strWhere

    
    strSQL = "" & _
    " Select 关联交易ID,卡类别ID,a.结算方式,a.交易流水号,a.交易说明, sum(原始金额) as 原始金额, sum(已退金额) as 已退金额, sum(原始金额)-sum(已退金额) as 剩余未退金额" & _
    " From (" & strSQL & ") A " & _
    " Group by a.关联交易ID,a.卡类别ID,a.结算方式,a.交易流水号,a.交易说明"
    
    Set rsSwapRecord_Out = zlDatabase.OpenSQLRecord(strSQL, "获取三方交易的原始金额及未退金额", str关联交易IDs, lng关联交易ID)
    zlGetThirdMoneyInforRecordFromSwapID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlMoveRowBalanceFromSwapID(ByVal vsBalance As VSFlexGrid, _
    lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal lng关联交易ID As Long, _
    ByVal lng预交ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据关联交易ID,删除对应的结算列表中对应的行
    '入参:vsBalance-网格
    '     lngCardTypeID-卡类别ID
    '     lng关联交易ID-关联交易ID
    '     lng预交ID-预交款记录的ID
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-30 11:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As clsBalanceItem
    Dim lngRow As Long
    On Error GoTo errHandle
    With vsBalance
        i = 1
        lngRow = .Row
        Do While i <= .Rows - 1
            If zlGetBalanceItemFromBalanceGrid(vsBalance, i, objItem) Then
                If objItem.关联交易ID = lng关联交易ID And objItem.预交ID = lng预交ID _
                    And objItem.卡类别ID = lngCardTypeID And objItem.消费卡 = bln消费卡 Then
                    .RowData(i) = ""
                    Set objItem = Nothing
                   '满足条件
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
Public Function zlMoveRowBalanceFromBalanceType(ByVal vsBalance As VSFlexGrid, int类型 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据byt类型,删除对应的结算列表中对应的行
    '入参:vsBalance-网格
    '     int类型-类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    '     lng关联交易ID-关联交易ID
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-30 11:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As clsBalanceItem
    Dim lngRow As Long, intTYPE As Integer
    On Error GoTo errHandle
    With vsBalance
        i = 1
        lngRow = .Row
        Do While i <= .Rows - 1
            intTYPE = Val(.TextMatrix(i, .ColIndex("类型")))
            If int类型 = intTYPE Then
                .RowData(i) = ""
                .RemoveItem i   '满足条件
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
    Optional lngNotRow As Long = -1, Optional bln含误差费 As Boolean = True, _
    Optional objCurItem As clsBalanceItem, Optional ByVal bln作废 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算结算信息
    '入参:objBalanceInfor-当前的结算信息
    '     objCurItem-当前结算信息项(未包含在列表中)
    '出参:objBalanceInfor-计算后的结算信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-11 17:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objItem As clsBalanceItem
    On Error GoTo errHandle
    
    
    objBalanceInfor.未付合计 = 0: objBalanceInfor.已付合计 = 0
    With vsBlance
        For i = 1 To .Rows - 1
            If i <> lngNotRow Then
                If zlGetBalanceItemFromBalanceGrid(vsBlance, i, objItem) Then
                    If (bln含误差费 And objItem.结算性质 = 9) Or objItem.结算性质 <> 9 Then
                        If Not (bln作废 And objItem.消费卡 And objItem.是否退款 And objItem.是否预交) Then
                            objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + objItem.结算金额, 6)
                        End If
                    End If
                End If
            End If
        Next
    End With
    If Not objCurItem Is Nothing Then
        objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + objCurItem.结算金额, 6)
    End If
    objBalanceInfor.未付合计 = RoundEx(objBalanceInfor.当前结帐 - objBalanceInfor.已付合计, 6)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Function zlGetReadFeeDetailFromBalanceID(ByVal lng结帐ID As Long, int病人来源 As Integer, bln作废 As Boolean, ByVal blnNOMoved As Boolean, ByRef rsDetail_out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID,获取费用明细信息
    '入参:lng结帐ID-结帐ID
    '     bln作废-是否作废记录
    '     blnNOMoved-是否数据转移
    '     int病人来源-1:门诊;2-住院;0 -门诊或住院
    '出参:rsDetail_out-费用明细集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-02 19:14:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strFormat As String
    
    On Error GoTo errHandle
    strFormat = "99999999990." & String(IIf(gbytDec < 0, 1, gbytDec), "9")
    Select Case int病人来源
    Case 1  '门诊
        strSQL = "" & _
        "   (   Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,0 as 主页ID,收据费目,婴儿费," & _
        "           Sum(结帐金额) As 结帐金额,发生时间,max(医嘱序号) as 医嘱序号,Max(是否保密) As 是否保密  " & vbCrLf & _
        "       From " & IIf(blnNOMoved, "H", "") & "门诊费用记录 A " & vbCrLf & _
        "       where A.结帐ID=[1]  " & vbCrLf & _
        "       Group By 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,收据费目,婴儿费,发生时间 " & vbCrLf & _
        "    ) A "
        'strSQL = IIf(mblnNOMoved, "H", "") & "门诊费用记录 A "
    Case 2  '住院
        strSQL = IIf(blnNOMoved, "H", "") & "住院费用记录 A"
    Case Else '门诊和住院
        strSQL = "" & _
        " (     Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,0 as 主页ID,收据费目,婴儿费,结帐金额,发生时间,医嘱序号,是否保密 " & vbCrLf & _
        "       From " & IIf(blnNOMoved, "H", "") & "门诊费用记录 A " & vbCrLf & _
        "       Where A.结帐ID=[1] " & vbCrLf & _
        "       Union ALL " & vbCrLf & _
        "       Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,主页ID,收据费目,婴儿费,结帐金额,发生时间,医嘱序号,是否保密 " & vbCrLf & _
        "       From " & IIf(blnNOMoved, "H", "") & "住院费用记录 A " & vbCrLf & _
        "       Where A.结帐ID=[1]  " & vbCrLf & _
        " )  A"
    End Select
    
    strSQL = _
    "   Select Decode(门诊标志,1,'门诊',4,'门诊',Decode(Nvl(A.主页ID,0),0,'','第'||Nvl(A.主页ID,0)||'次')) As 类型," & vbCrLf & _
    "         A.NO as 单据号,Nvl(B.名称,'未知') as 开单科室,decode(nvl(a.是否保密,0),1,'***',Nvl(E.名称,D.名称)) as 项目," & vbCrLf & _
             IIf(gTy_System_Para.byt药品名称显示 = 2, "decode(nvl(a.是否保密,0),1,'***',E1.名称) as 商品名,", "") & vbCrLf & _
    "       A.收据费目 as 费目,Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费," & vbCrLf & _
    "       ltrim(rtrim(To_Char(" & IIf(bln作废, "-1*", "") & "A.结帐金额,'" & strFormat & "'))) as 结帐金额," & vbCrLf & _
    "       To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 费用时间" & vbCrLf & _
    " From " & strSQL & ",部门表 B,收费项目目录 D,收费项目别名 E" & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & vbCrLf & _
    " Where A.开单部门ID=B.ID And A.收费细目ID=D.ID" & vbCrLf & _
    "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & vbCrLf & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & vbCrLf, "") & _
    "       And A.结帐ID=[1]" & vbCrLf & _
    " Order by 类型 Desc,费用时间 Desc,单据号 Desc,A.序号"
    Set rsDetail_out = zlDatabase.OpenSQLRecord(strSQL, "根据结帐ID获取结算费用明细数据", lng结帐ID)
    zlGetReadFeeDetailFromBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetExceptionBalanceData(ByVal bytType As Byte, ByRef dtStartDate As Date, _
    ByVal dtEndDate As Date, ByVal str操作员 As String, _
    rsErrData_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取异常的结帐数据
    '入参:bytType:0-异常的结帐记录;1-异常的结帐作废记录
    '     bytDateRange-日期范围:
    '出参:rsErrData_Out-返回异常的结算数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-07 19:35:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere  As String, strTable As String, strSQL As String
    Dim rsTemp As ADODB.Recordset, str合约病人IDs As String
    Dim cllFilter As Collection, cllPati As Collection
    Dim objPati As clsPatientInfo
    
    On Error GoTo errHandle
    strWhere = "  And A.收费时间 Between [1] And [2] And A.操作员姓名 = [3] And A.结算状态 = 1"
    If bytType = 0 Then
        strWhere = strWhere & " And A.记录状态 In (1,3)"
    Else
        strWhere = strWhere & "And A.记录状态 = 2"
    End If
    
    strTable = "" & _
    " Select A.ID ,1 as 住院标志,0 as 门诊标志,A.NO,A.实际票号,A.病人ID,A.主页ID,A.开始日期,A.结束日期," & _
    "       Max(A.记录状态) As 记录状态,Sum(B.结帐金额) As 结帐金额,A.操作员姓名,A.收费时间," & _
    "       A.中途结帐,A.原因 as 合约单位,A.结帐类型,Max(b.病人ID) As 费用病人ID,Max(b.标识号) As 标识号, " & _
    "       Max(b.姓名) As 姓名,Max(b.性别) As 性别,Max(b.年龄) As 年龄,Max(b.费别) As 费别" & _
    " From 病人结帐记录 A,住院费用记录 B" & _
    " Where A.ID=B.结帐ID " & strWhere & _
    " Group By A.ID ,A.NO,A.实际票号,A.病人ID,a.主页id,A.开始日期,A.结束日期,A.操作员姓名,A.收费时间," & _
    "   A.中途结帐,A.原因,A.结帐类型 "
    
    strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & _
        Replace(Replace(strTable, "住院费用记录", "门诊费用记录"), "1 as 住院标志,0 as 门诊标志", "0 as 住院标志,1 as 门诊标志")
    
    strSQL = "" & _
    " Select A.ID as 结帐ID,decode(住院标志,1,decode(门诊标志,1,3,2),1) As 标志," & _
    "        decode(A.结帐类型,1,'门诊结帐',2,'住院结帐','') As 结帐类型 , " & _
    "        Decode(D.险类,NULL,NULL,'√') as 医保,A.NO as 单据号,A.实际票号 As 票据号," & _
    "        Decode(A.病人ID,Null,' ',A.病人ID) As 病人ID," & _
    "        Decode(Nvl(A.结帐类型,0),2,' ',Decode(A.病人ID,Null,' ',a.标识号)) As 门诊号," & _
    "        Decode(Nvl(A.结帐类型,0),1,' ',Decode(A.病人ID,Null,' ',a.标识号)) As 住院号," & _
    "        Decode(A.病人ID,Null,A.合约单位,a.姓名) As 姓名," & _
    "        Decode(A.病人ID,Null,' ',a.性别) As 性别," & _
    "        Decode(A.病人ID,Null,' ',a.年龄) As 年龄," & _
    "        Decode(A.病人ID,Null,' ',a.费别) As 费别," & _
    "        To_Char(A.开始日期,'YYYY-MM-DD') As 开始日期,To_Char(A.结束日期,'YYYY-MM-DD') As 结束日期," & _
    "        To_Char(Decode(A.记录状态,2,-1,1) *A.结帐金额,'999999999" & gstrDec & "') as 结帐金额," & _
    "        A.操作员姓名 as 操作员,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间," & _
    "        Decode(Nvl(A.中途结帐,0),1,'√',' ') 中途结帐,A.记录状态 as 记录状态,A.费用病人ID" & _
    " From ( " & strTable & ") A,保险结算记录 D,人员表 N" & _
    " Where  A.操作员姓名=N.姓名 " & _
    "       And (N.站点='" & gstrNodeNo & "' Or N.站点 is Null) And A.id = D.记录ID(+) and D.性质(+)=2" & _
    " Order by 收费时间 Desc,单据号 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取结帐异常数据", dtStartDate, dtEndDate, UserInfo.姓名)
    
    Do While Not rsTemp.EOF
        If Val(NVL(rsTemp!病人ID)) = 0 And NVL(rsTemp!姓名) = "" Then
            If InStr("," & str合约病人IDs & ",", "," & rsTemp!费用病人ID & ",") = 0 Then
                str合约病人IDs = str合约病人IDs & "," & rsTemp!费用病人ID
            End If
        End If
        rsTemp.MoveNext
    Loop
    
    If str合约病人IDs = "" Then
        Set rsErrData_Out = rsTemp
    Else
        '取合约单位名称
        str合约病人IDs = Mid(str合约病人IDs, 2)
        If gobjSquare.objOneCardComLib.zlGetMultiPatiInforFromPatiID(str合约病人IDs, cllPati) = False Then Exit Function
        
        Set rsErrData_Out = zlDatabase.CopyNewRec(rsTemp)
        Do While Not rsErrData_Out.EOF
            If Val(NVL(rsErrData_Out!病人ID)) = 0 And NVL(rsErrData_Out!姓名) = "" Then
                    Set objPati = cllPati("_" & rsTemp!费用病人ID)
                    rsErrData_Out!姓名 = objPati.工作单位
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

Public Function zlCheck病人审核(lng病人ID As Long, lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断病人是否已审核
    '入参:lng病人ID-病人ID
    '     lng主页ID-主页ID
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-18 13:22:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPati As clsPatientInfo
    
    On Error GoTo ErrHandler
    Set objPati = New clsPatientInfo
    objPati.病人ID = lng病人ID
    objPati.主页ID = lng主页ID
    If zlGetPatiInfoByPage(objPati) = False Then Exit Function
    
    '49501
    If gTy_System_Para.byt病人审核方式 = 0 Then
        zlCheck病人审核 = (objPati.审核标志 >= 1)
    Else
        zlCheck病人审核 = (objPati.审核标志 > 1)
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
    '功能:检查病人是否审核
    '     objBalanceAllCons-当前条件
    '出参:strMessage-错误信息
    '     objBalanceAllCons-返回strUnAuditTime属性的住院次数
    '编制:刘兴洪
    '日期:2015-01-05 14:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnAll As Boolean, lng主页ID As Long, i As Long
    Dim varData As Variant
    On Error GoTo errHandle
    
    '门诊不进行检查
    If bytEditType = g_Ed_门诊结帐 Or objPati Is Nothing Then zlCheckPatiIsVerfy = True: Exit Function
    
    If InStr(strPrivs, ";未审核病人中途结帐;") > 0 Or InStr(strPrivs, ";未审核病人出院结帐;") > 0 Then zlCheckPatiIsVerfy = True: Exit Function
    If objPati.主页ID = 0 Then zlCheckPatiIsVerfy = True: Exit Function
    
    If CStr(objPati.主页ID) = objBalanceAllCons.strAllTime Then  '只有最后一次未结
        If objPati.审核标志 = 0 Then
            strMessage = "当前病人未审核，你不能对未审核的病人进行结帐。"
            Exit Function
        End If
        zlCheckPatiIsVerfy = True: Exit Function
    End If
    blnAll = True
    varData = Split(objBalanceAllCons.strAllTime, ",")
    For i = 0 To UBound(varData)
        lng主页ID = Val(varData(i))
        If lng主页ID <> 0 Then
            If Not zlCheck病人审核(objPati.病人ID, lng主页ID) Then
                 objBalanceAllCons.strUnAuditTime = objBalanceAllCons.strUnAuditTime & "," & lng主页ID
            Else
                blnAll = False
            End If
        Else
            blnAll = False
        End If
    Next
    If objBalanceAllCons.strUnAuditTime <> "" Then objBalanceAllCons.strUnAuditTime = Mid(objBalanceAllCons.strUnAuditTime, 2)
    If blnAll Then
        strMessage = "该病人所有住院费用都没有审核，不能进行结帐！"
        Exit Function
    End If
    zlCheckPatiIsVerfy = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCheckIsThirdSwapFromBalanceID(ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定次数的结帐是否存在三方交易
    '入参:lng结帐ID-结帐ID
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-20 16:09:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select 1  From 病人预交记录 where 结帐ID=[1] and 卡类别ID is not null and mod(记录性质,10)<>1 and rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否存在三方交易", lng结帐ID)
    zlCheckIsThirdSwapFromBalanceID = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCheckMulitInterfaceNumValied(ByVal vsBlance As VSFlexGrid, ByRef objCard As Card, objBalanceInfor As clsBalanceInfo, Optional bln预交 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是正同时存在三种以上接口(不含三种)
    '返回:不含两种以上接口的,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strIDs As String, strCardTypeIDs As String, strTemp As String, objItem As clsBalanceItem, objItems As clsBalanceItems
    Dim intMousePointer As Integer
    Dim intCount As Integer, i As Long, int性质 As Integer, str结算方式 As String
    Dim varData As Variant, strErrMsg As String

    On Error GoTo errHandle
    
    intMousePointer = Screen.MousePointer
     If objCard Is Nothing Then zlCheckMulitInterfaceNumValied = True: Exit Function

    If bln预交 Or objCard.接口序号 <= 0 Then zlCheckMulitInterfaceNumValied = True: Exit Function

    strErrMsg = ""
    
   '医保算一个接口
   If objBalanceInfor.objInsure.险类 <> 0 And objBalanceInfor.是否保存结帐单 Then intCount = intCount + 1: strErrMsg = strErrMsg & "医保结算:" & Format(objBalanceInfor.医保支付合计, gstrDec)
   
   strIDs = "": strCardTypeIDs = "," & objCard.接口序号 '把当前结算方式排除，重复使用结算方式已有检查
   With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            int性质 = Val(.TextMatrix(i, .ColIndex("类型")))
            If zlGetBalanceItemFromBalanceGrid(vsBlance, i, objItem) Then
                strTemp = objItem.卡类别ID & "," & objItem.关联交易ID
                If InStr("34", int性质) > 0 And InStr(strIDs & "|", "|" & strTemp & "|") = 0 And objItem.消费卡 = False And objItem.卡类别ID > 0 Then
                    If zlGetBalanceItemsFromVsBalanceGrid(vsBlance, objItem, objItems) = False Then Exit Function
                    strIDs = strIDs & "|" & strTemp
                    If InStr(strCardTypeIDs & ",", "," & objItem.卡类别ID & ",") = 0 Then
                        strCardTypeIDs = strCardTypeIDs & "," & objItem.卡类别ID
                        intCount = intCount + 1
                        
                        If objItem.objCard Is Nothing Then
                            strErrMsg = strErrMsg & vbCrLf & objItem.结算方式 & ":" & objItems.结算金额
                        Else
                            strErrMsg = strErrMsg & vbCrLf & objItem.objCard.名称 & ":" & objItems.结算金额
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    If intCount >= 3 Then
        Screen.MousePointer = 0
        Call MsgBox("注意:" & vbCrLf & "   本系统目前只支持三种以下接口,现在已经存在如下接口交易:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
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
    '功能:根据当前退款信息对象，检查退款是否合法
    '入参:objCurDelItem-当前退款项
    '     objThirdSwap-三方接口
    '出参:objItems_Out-当前退款信息列
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-29 10:16:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, objCard As Card, objItemsTemp As clsBalanceItems, objItemsPt As clsBalanceItems
    Dim bln消费卡 As Boolean, lngCardTypeID As Long, lngCurCardTypeID As Long
    Dim dblMoney As Double, dbl冲预交 As Double, dblDelMoney As Double, blnSingleDel As Boolean
    Dim i As Long, j As Long, intMousePointer As Integer, strErrMsg As String, strExpend As String, strDefaultBalance As String
    Dim blnFind As Boolean, blnAdd As Boolean, blnDelCash As Boolean
    
    On Error GoTo errHandle

      
    If objCurDelItem Is Nothing Then Exit Function
    Set objCard = objCurDelItem.objCard
    
    If objCard Is Nothing Then Exit Function
    
    lngCurCardTypeID = objCurDelItem.卡类别ID
    
    
    dblDelMoney = RoundEx(Abs(objCurDelItem.结算金额), 2)
    
    intMousePointer = Screen.MousePointer
    
    If objBalanceInfor.冲预交合计 = 0 Then
        Screen.MousePointer = 0
        MsgBox "当前无预交款退款，不能使用『" & objCard.名称 & "』进行退款操作！", vbInformation + vbOKOnly, gstrSysName
        Screen.MousePointer = intMousePointer
        Exit Function
    End If
    If dblDelMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox "『" & objCard.名称 & "』未输入退款金额，不能进行退款操作！", vbInformation + vbOKOnly, gstrSysName
        Screen.MousePointer = intMousePointer
        Exit Function
    End If
    
    blnSingleDel = objThirdSwap.zlThirdSwapIsSwapNOCall(lngCurCardTypeID, False, strErrMsg, strExpend)
    
    objCurDelItem.是否退款分交易 = blnSingleDel
    objCurDelItem.是否退款 = True
    objCurDelItem.是否预交 = True
    
    
    Set objItemsPt = New clsBalanceItems
    Set objItems_Out = New clsBalanceItems
    With vsDeposit
        For i = .Rows - 1 To 1 Step -1
            lngCardTypeID = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
            bln消费卡 = Val(.TextMatrix(i, .ColIndex("是否消费卡"))) = 1
            dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
            
            If lngCurCardTypeID = lngCardTypeID And bln消费卡 = False And Trim(.TextMatrix(i, .ColIndex("单据号"))) <> "" And dbl冲预交 > 0 Then
                If dblDelMoney = 0 Then Exit For
                If dblDelMoney >= dbl冲预交 Then
                    dblMoney = dbl冲预交
                    dblDelMoney = RoundEx(dblDelMoney - dbl冲预交, 2)
                Else
                    dblMoney = dblDelMoney
                    dblDelMoney = 0
                End If
                
                Set objItem = zlCopyNewItemFromBalanceItem(objCurDelItem)
                If objCard Is Nothing Then Set objCard = zlGetCardFromCardType(lngCardTypeID, False, Trim(.TextMatrix(i, .ColIndex("结算方式"))))
                
                objItem.结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式")))
                objItem.关联交易ID = Val(.TextMatrix(i, .ColIndex("关联交易ID")))
                objItem.卡类别ID = lngCardTypeID
                objItem.卡号 = Trim(.TextMatrix(i, .ColIndex("卡号")))
                objItem.交易流水号 = Trim(.TextMatrix(i, .ColIndex("交易流水号")))
                objItem.交易说明 = Trim(.TextMatrix(i, .ColIndex("交易说明")))
                objItem.结算号码 = Trim(.TextMatrix(i, .ColIndex("结算号码")))
                objItem.结算金额 = RoundEx(-1 * dblMoney, 6)
                objItem.结算摘要 = Trim(.TextMatrix(i, .ColIndex("摘要")))
                objItem.门诊结帐 = IIf(objBalanceInfor.结算类型 = 1, True, False)
                objItem.是否退款分交易 = blnSingleDel
                objItem.冲销ID = objBalanceInfor.冲销ID
                objItem.结算IDs = objBalanceInfor.结帐ID
                objItem.结帐ID = objBalanceInfor.结帐ID
                objItem.结帐时间 = objBalanceInfor.结帐时间
                objItem.结算性质 = objCard.结算性质
                objItem.是否预交 = True
                objItem.是否退款 = True
                objItem.结算类型 = 3 '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                objItem.预交ID = Val(.TextMatrix(i, .ColIndex("预交ID")))
                objItem.是否密文 = objCard.卡号密文规则 <> ""
                objItem.校对标志 = 1
                
                Set objItem.objCard = objCard
                
                Set objItemsTemp = New clsBalanceItems
                objItemsTemp.AddItem objItem
                objItemsTemp.结算金额 = objItem.结算金额
                objItemsTemp.收费类型 = 1
                blnAdd = False
                
                 If Not objThirdSwap.zlThirdReturnCashCheck(objCard, objItemsTemp, blnDelCash, strDefaultBalance) Then
                    '1.禁止退现
                    objItem.是否允许退现 = False
                    objItem.是否强制退现 = blnDelCash
                    objItem.是否允许删除 = objItem.是否强制退现
                    blnAdd = True
                 Else
                    If blnDelCash = False Then  '是否缺省退现
                          '允许退现，可以删除
                          objItem.是否允许编辑 = False
                          objItem.是否允许删除 = True
                          objItem.是否强制退现 = True
                          objItem.是否允许退现 = True: blnAdd = True
                      ElseIf strDefaultBalance <> "" Then
                          blnFind = False
                          For j = 1 To objItemsPt.Count
                              If objItemsPt(j).结算方式 = strDefaultBalance Then
                                  objItemsPt(j).结算金额 = objItemsPt(j).结算金额 + objItem.结算金额
                                  objItemsPt.结算金额 = objItemsPt.结算金额 + objItem.结算金额
                                  blnFind = True
                                  Exit For
                              End If
                          Next
                          If Not blnFind Then
                              Set objItem = New clsBalanceItem
                              With objItem
                                  Set .objCard = zlGetCardFromBalanceName(strDefaultBalance)
                                  .结算方式 = strDefaultBalance
                                  .结算金额 = RoundEx(-1 * dblMoney, 6)
                                  .是否退款 = True
                                  .是否允许编辑 = False
                                  .是否允许删除 = True
                                  .结算性质 = .objCard.结算性质
                                  .结算类型 = 0 '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                                  .Tag = "指定预交退款"
                              End With
                              objItemsPt.AddItem objItem
                              objItemsPt.结算金额 = RoundEx(objItemsPt.结算金额 + objItem.结算金额, 6)
                          End If
                      End If
                 End If
                 
                 If blnAdd Then  '未找到，需要增加
                    objItems_Out.AddItem objItem
                    objItems_Out.结算金额 = RoundEx(objItems_Out.结算金额 + objItem.结算金额, 2)
                 End If
            End If
        Next
    End With
    If objItems_Out.Count <> 0 And Not blnSingleDel Then
        '多交易一次退款
        Set objItem = zlCopyNewItemFromBalanceItem(objCurDelItem)
        If objCard Is Nothing Then Set objCard = zlGetCardFromCardType(lngCardTypeID, False, objCurDelItem.结算方式)
        objItem.是否退款 = True
        For i = 1 To objItems_Out.Count  '只要一有项满足打条件的，就允许主项有对应的处理
            If objItems_Out(i).是否允许删除 Then objItem.是否允许删除 = True
            If objItems_Out(i).是否强制退现 Then objItem.是否强制退现 = True
            If objItems_Out(i).是否允许退现 Then objItem.是否允许退现 = True
        Next
        
        Set objItem.objTag = objItems_Out
        objItem.结算金额 = objItems_Out.结算金额
        Set objItems_Out = New clsBalanceItems
        objItems_Out.AddItem objItem
        objItems_Out.结算金额 = objItems_Out.结算金额 + objItem.结算金额
    End If
        
    '加上普通的结算方式
    For Each objItem In objItemsPt
        objItems_Out.AddItem objItem
        objItems_Out.结算金额 = objItems_Out.结算金额 + objItem.结算金额
    Next
    
    If objItems_Out.结算金额 = 0 Then
        Screen.MousePointer = 0
        MsgBox "在预交款结算结算表中，不存在『" & objCard.名称 & "』的预交款，不能用该结算信息进行退款操作！", vbInformation + vbOKOnly, gstrSysName
        Screen.MousePointer = intMousePointer
        Set objItems_Out = Nothing
        Exit Function
    End If
    
    If RoundEx(Abs(objItems_Out.结算金额), 2) < RoundEx(Abs(objCurDelItem.结算金额), 2) Then
        Screen.MousePointer = 0
        MsgBox "在预交款结算结算表中，『" & objCard.名称 & "』的原始结算金额小于了本次退款金额，不能用该结算信息进行退款操作！" & vbCrLf & _
               "原始结算金额:" & Format(objItems_Out.结算金额, "0.00") & vbCrLf & _
               "本次退款金额:" & Format(objCurDelItem.结算金额, "0.00"), vbInformation + vbOKOnly, gstrSysName
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
    '功能:设置编辑状态
    '入参:blnSetRowData-是否将objItem设置给Rowdata属性
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-09-03 10:06:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    If objItem Is Nothing Then Exit Sub
    
    lngRow = objItem.行号
    With vsBlance
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        .TextMatrix(lngRow, .ColIndex("结算状态")) = IIf(objItem.是否结算, 1, 0)
        .TextMatrix(lngRow, .ColIndex("编辑状态")) = IIf(objItem.是否允许编辑, 1, 0) & "|" & IIf(objItem.是否允许删除, 1, 0)
        If blnSetRowData Then .RowData(lngRow) = objItem
    End With
End Sub




Public Function zlGetItemFromVsDepositRow(ByVal vsDeposit As VSFlexGrid, ByVal lngRow As Long, ByRef objItem_Out As clsBalanceItem) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定预交款行，获取该行的Item对象
    '入参:vsDeposit-预交款列表
    '出参:objItem_Out-获取指定信息项
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-06-14 15:14:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngCardTypeID As Long, bln消费卡 As Boolean, dblMoney As Double
    
    Err = 0: On Error GoTo errHandle:
    Set objItem_Out = New clsBalanceItem
    '计算三方退款信息或转帐信息
    With vsDeposit
        dblMoney = Val(.TextMatrix(lngRow, .ColIndex("冲预交")))
        lngCardTypeID = Val(.TextMatrix(lngRow, .ColIndex("卡类别ID")))
        bln消费卡 = Val(.TextMatrix(lngRow, .ColIndex("是否消费卡"))) = 1
        
        Set objItem_Out.objCard = zlGetCardFromCardType(lngCardTypeID, bln消费卡, Trim(.TextMatrix(lngRow, .ColIndex("结算方式"))))
        objItem_Out.是否转帐 = Val(.TextMatrix(lngRow, .ColIndex("是否转帐及代扣"))) = 1
        objItem_Out.结算性质 = Val(.TextMatrix(lngRow, .ColIndex("结算性质")))
        objItem_Out.结算IDs = ""
        objItem_Out.交易流水号 = Trim(.TextMatrix(lngRow, .ColIndex("交易流水号")))
        objItem_Out.交易说明 = Trim(.TextMatrix(lngRow, .ColIndex("交易说明")))
        objItem_Out.卡号 = Trim(.TextMatrix(lngRow, .ColIndex("卡号")))
        objItem_Out.关联交易ID = Val(.TextMatrix(lngRow, .ColIndex("关联交易ID")))
        objItem_Out.结算金额 = Val(.TextMatrix(lngRow, .ColIndex("冲预交")))
        objItem_Out.结算方式 = Trim(.TextMatrix(lngRow, .ColIndex("结算方式")))
        objItem_Out.结算号码 = Trim(.TextMatrix(lngRow, .ColIndex("结算号码")))
        objItem_Out.结算摘要 = Trim(.TextMatrix(lngRow, .ColIndex("摘要")))
        If lngCardTypeID <> 0 Then
            objItem_Out.结算类型 = IIf(Not bln消费卡, 3, 5)    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        ElseIf objItem_Out.结算性质 = 7 Then
            objItem_Out.结算类型 = 4
        Else
            objItem_Out.结算类型 = 0 '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        End If
        objItem_Out.是否预交 = True
        If .ColIndex("余额") >= 0 Then
            objItem_Out.未退金额 = Val(.TextMatrix(lngRow, .ColIndex("余额")))
        End If
        If .ColIndex("金额") >= 0 Then objItem_Out.原始金额 = Val(.TextMatrix(lngRow, .ColIndex("金额")))
        objItem_Out.卡类别ID = lngCardTypeID
        objItem_Out.消费卡 = bln消费卡
        objItem_Out.是否密文 = IIf(objItem_Out.objCard.卡号密文规则 <> "", True, False)
        objItem_Out.结帐时间 = CDate(.TextMatrix(lngRow, .ColIndex("收款日期")))
        objItem_Out.是否允许退现 = objItem_Out.objCard.是否退现
        objItem_Out.预交ID = Val(.TextMatrix(lngRow, .ColIndex("预交ID")))
        
        objItem_Out.消费卡ID = 0
        objItem_Out.是否退款分交易 = True
        objItem_Out.是否退款 = False
        objItem_Out.是否允许编辑 = False
        objItem_Out.是否允许删除 = True
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
    '功能:将所有项目的结算金额取返数
    '入参:objItems-项目集
    '出参:objItems-返回的返数集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-09-05 12:02:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long

    On Error GoTo errHandle
    
    If objCurItems Is Nothing Then Exit Sub
    objCurItems.结算金额 = RoundEx(-1 * objCurItems.结算金额, 6)
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
    '功能:将所有项目的结算金额取返数
    '入参:objCurItem-当前项目集
    '出参:objCurItem-返回的返数集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-09-05 12:02:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItems As clsBalanceItems

    On Error GoTo errHandle
    If objCurItem Is Nothing Then Exit Sub
    
    objCurItem.结算金额 = RoundEx(-1 * objCurItem.结算金额, 6)
    Set objItems = objCurItem.objTag
    If objItems Is Nothing Then Exit Sub
    objItems.结算金额 = RoundEx(-1 * objItems.结算金额, 6)
    For i = 1 To objItems.Count
        objItems(i).结算金额 = RoundEx(-1 * objItems(i).结算金额, 6)
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlGetFromIDToBalanceData(ByVal lng结帐ID As Long, ByVal blnNOMoved As Boolean, _
    ByRef rsOutBalance As ADODB.Recordset, Optional blnView As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID来获取结算数据
    '入参:lng结帐ID-结帐ID
    '     blnNoMoved-是否已经转移到后备表中
    '     blnView-是否查阅
    '出参:rsOutBalance-结帐数据
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 15:32:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, objCard As Card
    Dim rsNew  As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim blnExistOfflineYb As Boolean
    
    On Error GoTo errHandle
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡;6-误差费
    strSQL = "" & _
    "   Select  A.ID, " & _
    "        Case when Mod(A.记录性质,10)=1 then 1  " & _
    "             when (nvl(M.性质,0)=3 or nvl(M.性质,0)=4) and nvl(a.卡类别ID,0)=0  then 2 " & _
    "             when nvl(A.卡类别ID,0)<>0  then  3 " & _
    "             when nvl(M.性质,0)=9 then 6 " & _
    "             else 0 end as 类型, " & _
    "        Mod(A.记录性质,10) as 记录性质,A.结算方式,A.冲预交,A.摘要,A.卡类别ID,A.结算卡序号,A.结算号码,A.卡号,A.交易流水号," & vbNewLine & _
    "        nvl(C1.自制卡,0) as 自制卡, nvl(C1.是否退现,0) as 是否退现," & vbNewLine & _
    "        nvl(C1.是否全退,0) as 是否全退, 0 as 是否转帐及代扣," & vbNewLine & _
    "        Decode(C1.是否密文,NULL,0,1) as 是否密文,C1.名称  as 卡类别名称," & vbNewLine & _
    "        A.交易说明,A.结算序号,A.校对标志,decode(nvl(M.性质,0),3,1,4,1,0) as 医保,0 as 消费卡id," & _
    "        nvl(M.性质,0) as 结算性质,A.ID as 预交ID,A.关联交易ID,nvl(a.附加标志,0) as 附加标志" & vbNewLine & _
    "   From  病人预交记录 A ,结算方式 M,消费卡类别目录 C1" & _
    "   Where A.结帐ID= [1] And A.结算方式=M.名称(+) And a.结算卡序号 = c1.编号(+)  " & _
    "         And ( nvl(A.结算卡序号,0)=0 OR  Mod(a.记录性质,10) =1) "
    
    If Not blnView Then
        '--门诊费用转住院时，一次结算（医疗卡支付）的门诊费用通过多次转入后产生了多个住院预交单据，这些预交单据的关联交易ID相同
        '--在结帐预交退款时，关联交易ID相同的记录在预交记录中只有一条，三方退款信息中有多条，所以，
        '   先排除预交记录中的数据，并将三方退款信息中的金额按结算方式为NULL处理，三方退款信息在外面单独处理
        '--存在预交退款时，就不存在非预交金额退款，所以病人预交记录与三方退款信息的关联条件可以为(结帐id,卡类别id）
        strSQL = strSQL & _
        "         And Not Exists (Select 1 From 三方退款信息" & _
        "                         Where 结帐id = a.结帐id And 卡类别id = a.卡类别id And a.记录性质 = 2 And a.冲预交 < 0)"

        strSQL = strSQL & " Union ALL " & _
        " Select a.Id, 0 As 类型, Mod(a.记录性质, 10) As 记录性质, '' As 结算方式, a.冲预交, '' As 摘要," & _
        "       Null As 卡类别id, Null As 结算卡序号, '' As 结算号码, '' As 卡号, '' As 交易流水号, " & _
        "       0 As 自制卡, 0 As 是否退现, 0 As 是否全退, 0 As 是否转帐及代扣, 0 As 是否密文, '' As 卡类别名称," & _
        "        '' As 交易说明, a.结算序号, 0 As 校对标志, 0 As 医保, 0 As 消费卡id, 0 As 结算性质," & _
        "       a.Id As 预交id, Null As 关联交易id, 0 As 附加标志" & _
        " From 病人预交记录 A" & _
        " Where a.结帐id = [1] And Mod(a.记录性质, 10) <> 1 And a.卡类别id Is Not Null" & _
        "       And Exists (Select 1 From 三方退款信息" & _
        "                   Where 结帐id = a.结帐id And 卡类别id = a.卡类别id And a.记录性质 = 2 And a.冲预交 < 0)"
    End If
    
    strSQL = strSQL & " Union ALL " & _
    "   Select A.ID,5 as  类型,Mod(A.记录性质,10) as 记录性质,A.结算方式,-1*nvl(b.应收金额,0) as 冲预交,A.摘要,A.卡类别ID,A.结算卡序号," & _
    "        A.结算号码,B.卡号,B.交易流水号,nvl( M.自制卡,0) as 自制卡, " & _
    "        nvl( M.是否退现,0) as 是否退现,nvl(M.是否全退,0) as 是否全退,0 as 是否转帐及代扣," & _
    "        nvl(M.是否密文,0) as  是否密文," & _
    "        M.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志,0 as 医保,B.消费卡id,M1.性质 as 结算性质,A.ID as 预交ID,A.关联交易ID,nvl(a.附加标志,0) as 附加标志 " & _
    "   From 病人预交记录 A ,病人卡结算记录 B,消费卡类别目录 M,结算方式 M1 " & _
    "   Where  a.Id = b.结算Id  and a.结算卡序号 = m.编号 And A.结算方式=M1.名称(+) And A.结帐ID = [1] and Mod(A.记录性质,10)<>1 "
       
    strSQL = "" & _
    " Select a.类型, a.记录性质, a.结算方式, a.摘要, a.卡类别id, a.卡类别名称, a.自制卡, a.结算卡序号," & _
    "        a.结算号码, a.卡号, a.交易流水号, A. 交易说明, a.结算序号, a.校对标志, a.医保, a.消费卡id," & _
    "        a.是否密文, a.是否全退, a.是否转帐及代扣, a.是否退现, Nvl(a.冲预交, 0) As 冲预交," & _
    "        a.结算性质 As 性质, a.预交id, a.关联交易id, a.附加标志" & _
    " From (" & strSQL & ") A" & _
    " Order by 类型"
    
    If blnNOMoved Then
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
        strSQL = Replace(strSQL, "病人卡结算记录", "H病人卡结算记录")
        strSQL = Replace(strSQL, "三方退款信息", "H三方退款信息")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取结帐数据", lng结帐ID)
    blnExistOfflineYb = False
    rsTemp.Filter = "卡类别ID<>0"
    If rsTemp.RecordCount > 0 Then
        rsTemp.Filter = ""
        Set rsNew = zlDatabase.CopyNewRec(rsTemp)
        
        rsNew.Filter = "卡类别ID<>0"
        Do While Not rsNew.EOF
            If ZlGetPayCard(Val(NVL(rsNew!卡类别ID)), objCard) Then
                rsNew!卡类别名称 = objCard.名称
                rsNew!自制卡 = IIf(objCard.自制卡, 1, 0)
                rsNew!是否密文 = IIf(objCard.卡号密文规则 = "", 0, 1)
                rsNew!是否全退 = IIf(objCard.是否全退, 1, 0)
                rsNew!是否退现 = IIf(objCard.是否退现, 1, 0)
                rsNew!是否转帐及代扣 = IIf(objCard.是否转帐及代扣, 1, 0)
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
Public Function zlGetOldOfflineBalanceFromBlanaceID(ByVal lng冲销ID As Long, ByVal objBalanceInfor As clsBalanceInfo, _
    ByVal strOffLineBalances As String, ByRef objBalanceItems As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取原始的脱机医保结算集
    '入参:lng冲销id
    '     strOffLineBalances-脱机结算方式,格式：结算方式|结算方式...
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-12-23 19:32:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL  As String
    Dim objItem As clsBalanceItem
    Dim objCard As Card
    
    On Error GoTo errHandle
    
    Set objBalanceItems = New clsBalanceItems
    If strOffLineBalances = "" Then Exit Function
    
    strSQL = "" & _
    "   Select 结算方式,mod(记录性质,10) as 记录性质,冲预交 From 病人预交记录  " & _
    "   Where 结帐ID IN(Select distinct B.ID From 病人结帐记录 A,病人结帐记录 B " & _
    "                   Where a.ID=[1] And  A.NO=B.NO And B.记录状态 in (1,3)) And Mod(记录性质,10)<>1 And instr([2] ,'|'|| 结算方式||'|')>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取原始结帐金额-脱机医保", lng冲销ID, "|" & strOffLineBalances & "|")
    Do While Not rsTemp.EOF
        Set objItem = New clsBalanceItem
        Set objCard = zlGetCardFromCardType(0, False, NVL(rsTemp!结算方式))
       
        objItem.结算方式 = NVL(rsTemp!结算方式)
        objItem.结帐ID = objBalanceInfor.结帐ID
        objItem.结算IDs = objBalanceInfor.结帐ID
        objItem.冲销ID = objBalanceInfor.冲销ID
        objItem.结帐时间 = objBalanceInfor.结帐时间
        objItem.结算类型 = 0
        objItem.结算金额 = RoundEx(rsTemp!冲预交, 6)
        objItem.是否允许删除 = True
        objItem.是否允许退现 = True
        objItem.是否允许编辑 = False
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
Public Function zlGetLedDisplayBankDatasFromVsBalance(ByVal vsBlance As VSFlexGrid, ByRef cllBanks_out As Collection, ByVal dbl医保帐户余额 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算列表，获取显示在Led上的结算数据集
    '入参:objPati-病人信息集
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-09-26 16:21:26
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
            '医保交易
            If .TextMatrix(i, .ColIndex("结算方式")) <> "" Then
                '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                Select Case Val(.TextMatrix(i, .ColIndex("类型")))
                Case 2 '医保
                    cllYBBanks.Add Array(.TextMatrix(i, .ColIndex("结算方式")) & ":", Format(Val(.TextMatrix(i, .ColIndex("结算金额"))), "0.00"))
                Case 3 '三方接口交易
                    cllThirdBanks.Add Array(.TextMatrix(i, .ColIndex("结算方式")) & ":", Format(Val(.TextMatrix(i, .ColIndex("结算金额"))), "0.00"))
                Case 4 ' 一卡通交易
                    cllOldCardOneBanks.Add Array(.TextMatrix(i, .ColIndex("结算方式")) & ":", Format(Val(.TextMatrix(i, .ColIndex("结算金额"))), "0.00"))
                Case Else
                    cllPTBanks.Add Array(.TextMatrix(i, .ColIndex("结算方式")) & ":", Format(Val(.TextMatrix(i, .ColIndex("结算金额"))), "0.00"))
                End Select
            End If
        Next
    End With
    
    Set cllBanks_out = New Collection
    
    If cllYBBanks.Count <> 0 Then
        cllBanks_out.Add Array("医保结算:", Format(dbl医保帐户余额, "0.00"))
        For i = 1 To cllYBBanks.Count
            cllBanks_out.Add cllYBBanks(i)
        Next
    End If
    
    Set cllYBBanks = Nothing
    
    If cllThirdBanks.Count <> 0 Then
        cllBanks_out.Add Array("一卡通结算:", "")
        For i = 1 To cllThirdBanks.Count
            cllBanks_out.Add cllThirdBanks(i)
        Next
    End If
    Set cllThirdBanks = Nothing
    
    If cllOldCardOneBanks.Count <> 0 Then
        cllBanks_out.Add Array("一卡通结算(老):", "")
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
    '功能:根据当前选择的诊断获取因取诊断ID
    '入参:intComboxIdex-索引
    '     cboDiag-诊断下拉框
    '返回:医嘱ID
    '编制:刘兴洪
    '日期:2019-01-24 11:22:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str医嘱ID As String
    On Error GoTo errHandle
    If intComboxIdex < 0 Or intComboxIdex > cboDiag.ListCount - 1 Or cboDiag.Tag = "" Then Exit Function
    zlGetDiagIDFromComboxDiag = Split(cboDiag.Tag & ",,,", ",")(intComboxIdex)
    Exit Function
errHandle:
    zlGetDiagIDFromComboxDiag = ""
End Function
Public Sub zlLoadDiagnosDataToCombox(ByVal frmMain As Object, ByVal rsAllDiagnos As ADODB.Recordset, ByRef cboDiag As ComboBox)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加诊断信息给Combox控件
    '入参:str诊断IDs-医嘱IDs
    '     cboDiag-加载诊断的下拉控件
    '     rsAllDiagnos-诊断的记录集
    '编制:刘兴洪
    '日期:2019-01-23 18:03:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str诊断ID As String, strTemp As String
    Dim lngWidth As Single, lngTemp As Single
    Dim j As Long
    On Error GoTo errHandle
    
  
    str诊断ID = zlGetDiagIDFromComboxDiag(cboDiag.ListIndex, cboDiag)
    cboDiag.Clear
    cboDiag.AddItem "所有诊断"
    cboDiag.ListIndex = cboDiag.NewIndex
    cboDiag.Tag = "0"
    
    If rsAllDiagnos Is Nothing Then Exit Sub
    If rsAllDiagnos.State <> 1 Then Exit Sub
    lngWidth = cboDiag.Width
    
    With rsAllDiagnos
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            strTemp = zlFormatID(!诊断ID)
            If InStr("," & cboDiag.Tag & ",", "," & strTemp & ",") = 0 Then
                cboDiag.Tag = cboDiag.Tag & "," & strTemp
                cboDiag.AddItem NVL(!诊断描述)
                
                j = frmMain.TextWidth("L") + 15
                If j * zlCommFun.ActualLen(NVL(!诊断描述)) > 6465 Then
                    lngTemp = 6465
                Else
                    lngTemp = j * zlCommFun.ActualLen(NVL(!诊断描述)) + frmMain.TextWidth("刘") * 3
                End If
                If lngWidth < lngTemp Then lngWidth = lngTemp
                If strTemp = str诊断ID Then cboDiag.ListIndex = cboDiag.NewIndex
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

Public Function zlGetSelfPaymentMode(ByVal lngModule As Long, ByVal lng病人ID As Long, ByVal str主页IDS As String, _
    ByVal lng结帐ID As Long, ByVal rsFeelists As ADODB.Recordset, _
    ByRef strSelfPaymentMode_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取自费方式
    '入参:lngModule-调用模块号
    '     lng病人ID-病人ID
    '     str主页Ids-当前主页IDs,多个用逗号分离
    '     lng结帐ID-如果是医保病人，在医保结算后，调用本接口，传入当前的结帐ID,
    '               如果是普通病人，在读取费用明细后，调用本接口，结帐ID为0,但传入了rsFeeLists记录集
    '     rsFeeLists-传入的本次结帐的费用明细数据（住院,单据号 ,项目,费目,ID,序号,记录性质,记录状态,执行状态,主页ID,计算单位,
    '    数量, 价格,应收金额,实收金额,未结金额,结帐金额, 统筹金额,类型, 收费类别,收费类别名,费别, 婴儿费, 执行部门id,
    '    科室,开单部门ID,开单人, 保险大类id,收费细目ID,门诊标志,排序, 医嘱序号,时间,登记时间）(即传入现结帐程序中的费用明细集）
    '
    '出参:strSelfPaymentMode_Out-返回结算方式,本接口返回true时有效,格式为:结算方式1,结算金额1,结算号码1||结算方式2,结算金额2,结算号码2||....
    '返回:获取返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-05-05 14:41:27
    '调用时机:
    '     1.普通病人:在读取病人费用明细后调用本接口
    '     2.医保病人：进行医保结算后，调用本接口
    '应用场景：慈善救助
    '    1、慈善救助得患者需病区得护士进行身份的核对，患者在结帐的时候不会主动说自己是慈善救助的对象，所以让结账员手工录入的话，容易漏掉这部分。
    '    2、慈善救助金额算法是：如果是医保病人，需要在医保结算完成之后才按照自费金额的xx%计算的，普通病人则直接是自费费用的xx%。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetSelfPaymentMode = False
End Function

Public Function zlAddfinancialTrancsToBalanceList(ByRef vsBlance As VSFlexGrid, ByVal dblMoney As Double, Optional dblTranMoney_out As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加转帐结算信息
    '入参:vsBlance-结算列表
    '     dblMoney-结算退款金额
    '出参:dblTranMoney_out-转账金额
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-09-09 10:57:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str结算号码 As String, str结算摘要 As String, str卡号 As String
    
    On Error GoTo errHandle
    dblTranMoney_out = 0
    If Not (gTy_System_Para.TY_Balance.bln预交退指定结算方式 And gTy_System_Para.TY_Balance.str预交退款结算方式 <> "") Or dblMoney >= 0 Then
        zlAddfinancialTrancsToBalanceList = True: Exit Function
    End If
   
    Call ClearBalanceList(vsBlance, gTy_System_Para.TY_Balance.str预交退款结算方式)
    With vsBlance
        If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) <> "" Or .Rows <= 1 Then .Rows = .Rows + 1
        i = .Rows - 1
        
        Call zlPlugin_GetfinancialTrancsBalanceInfor(gTy_System_Para.TY_Balance.str预交退款结算方式, dblMoney, str结算号码, str结算摘要, str卡号)
        
        .RowData(.Rows - 1) = ""
        .TextMatrix(i, .ColIndex("类型")) = 0    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        .TextMatrix(i, .ColIndex("卡类别ID")) = 0
        .TextMatrix(i, .ColIndex("消费卡ID")) = 0
        .TextMatrix(i, .ColIndex("结算性质")) = 2    ''1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
        .TextMatrix(i, .ColIndex("编辑状态")) = 0   ''0-禁止删除;1-允许编辑金额;2-仅允许删除;3-允许删除及修改金额,4-禁止删除且禁止修改等等
        .TextMatrix(i, .ColIndex("结算状态")) = 0  '是否已结算:1-已结算;0-未结算
        .TextMatrix(i, .ColIndex("是否退现")) = 0
        .TextMatrix(i, .ColIndex("是否全退")) = 0
        .TextMatrix(i, .ColIndex("校对标志")) = 0
        .TextMatrix(i, .ColIndex("是否转账")) = 0
        .TextMatrix(i, .ColIndex("是否密文")) = 0
        .TextMatrix(i, .ColIndex("卡类别名称")) = ""
        .TextMatrix(i, .ColIndex("结算方式")) = gTy_System_Para.TY_Balance.str预交退款结算方式
        .TextMatrix(i, .ColIndex("结算金额")) = Format(dblMoney, "0.00")
        .TextMatrix(i, .ColIndex("结算号码")) = str结算号码
        .TextMatrix(i, .ColIndex("备注")) = str结算摘要
        .TextMatrix(i, .ColIndex("交易流水号")) = ""
        .TextMatrix(i, .ColIndex("交易说明")) = ""
        .TextMatrix(i, .ColIndex("卡号")) = str卡号
        .Cell(flexcpData, i, .ColIndex("卡号")) = str卡号
        dblTranMoney_out = Format(dblMoney, "0.00")
    End With
    zlAddfinancialTrancsToBalanceList = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetDepoistRowFromBalanceList(ByVal vsBlance As VSFlexGrid, ByVal str结算方式 As String) As Integer
    '功能：从结算列表中获取“剩余预交款按某种结算方式转账”列
    '入参：vsBlance-传入的结算信息列表
    '      ：str结算方式-结帐剩余款存为预交款的结算方式
    Dim i As Integer
    
    If vsBlance.Rows <= 1 Then Exit Function
    With vsBlance
        If .ColIndex("结算方式") = -1 Or .ColIndex("备注") = -1 Then Exit Function
        For i = 1 To vsBlance.Rows - 1
            If .TextMatrix(i, .ColIndex("结算方式")) = str结算方式 And Val(.TextMatrix(i, .ColIndex("类型"))) = 0 Then
                GetDepoistRowFromBalanceList = i: Exit For
            End If
        Next
    End With
End Function

Public Sub ClearBalanceList(ByVal vsBlance As VSFlexGrid, ByVal str结算方式 As String)
    '功能：预交款不足以结帐时，清除“结帐剩余款存预交”列
    '入参：vsBlance-传入的结算信息列表
    '      ：str结算方式-结帐剩余款存为预交款的结算方式
    Dim i As Integer
    
    If vsBlance.Rows <= 1 Then Exit Sub
    i = GetDepoistRowFromBalanceList(vsBlance, str结算方式)
    If i > 0 Then vsBlance.RemoveItem i

End Sub

Public Function zlCheckOtherSessionDoing(ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前结算是否正在被其它会话处理
    '入参:lng结帐ID-指定的结算序号
    '出参:
    '返回:被其他会话站用返回true,否则返回False
    '说明："病人预交记录.会话号"格式：V$session.SID+'_'+V$session.SERIAL#
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If lng结帐ID = 0 Then zlCheckOtherSessionDoing = False: Exit Function
    
    strSQL = "Select 1" & vbNewLine & _
            " From 病人预交记录 A, V$session B" & vbNewLine & _
            " Where a.会话号 = b.Sid || '_' || b.Serial# And a.结帐ID = [1] " & vbNewLine & _
            "       And b.Username Is Not Null And b.Audsid <> Userenv('sessionid')" & vbNewLine & _
            "       And Upper(b.Status) In ('ACTIVE', 'INACTIVE') And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查当前结算是否正在被其它会话处理", lng结帐ID)
    zlCheckOtherSessionDoing = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckBalanceOverFromBalanceID(ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID,检查当前结帐是否已经完成
    '返回:结帐完成返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-09-17 10:53:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "Select 1 From 病人结帐记录 where id=[1] And nvl(结算状态,0)=0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查当前结帐状态", lng结帐ID)
    zlCheckBalanceOverFromBalanceID = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlErrBalanceCheckFromPatiID(ByVal lng病人ID As Long, ByRef strErrNo_Out As String, ByRef blnDel_Out As Boolean, _
    Optional ByVal strCheckNO As String, Optional ByVal str当前操作 As String = "结帐") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID或当前的结帐NO，判断是否存在异常的结帐或退费单据
    '入参:lng病人ID-病人ID
    '     strCheckNo-不为空时，按结帐单据号进行检查
    '出参:
    '    strErrNo_Out-函数返回true时，表示返回的异常单据NO,返回False时，本参数为空
    '    blnDel_Out-当前返回的异单据是否为异常作废结帐单据
    '返回:0-不存在异常单据
    '     1-存在异常单据，当选择为继续结帐
    '     2-存在异常单据，当前选择为终止结帐
    '     3-存在异常单据，当前选择针对异常单据进行重结或重退
    '编制:刘兴洪
    '日期:2019-09-17 10:53:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lng结帐ID As Long, str操作员姓名 As String, strTittle As String
 
     
    strErrNo_Out = "": blnDel_Out = False
     
    strSQL = " " & _
    "    Select  a.No, a.ID, a.操作员姓名, decode(记录状态,2,2,1) As 异常类型,A.收费时间 " & _
    "    From 病人结帐记录 A" & _
    "    Where nvl(结算状态,0) = 1" & IIf(strCheckNO = "", " And 病人ID=[1]", " And No=[2]")
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查异常的结帐单据", lng病人ID, strCheckNO)
    
    If rsTemp.EOF Then zlErrBalanceCheckFromPatiID = 0: Exit Function   '0-不存在异常单据
    
    With rsTemp
        '理论上存在多个异常，需要重复检查
        Do While Not .EOF
            If zlCheckOtherSessionDoing(Val(NVL(rsTemp!ID))) Then
                MsgBox "注意:" & vbCrLf & _
                "    该病人存在异常的" & IIf(Val(NVL(rsTemp!异常类型)) = 2, "重退", "结帐") & "单据(" & NVL(rsTemp!NO) & ")且正在其他结帐窗口进行操作 ,你现在不能进行" & str当前操作 & "操作!", vbInformation + vbOKOnly, gstrSysName
                zlErrBalanceCheckFromPatiID = 2
                Exit Function
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    

    strErrNo_Out = NVL(rsTemp!NO): lng结帐ID = Val(NVL(rsTemp!ID))
    blnDel_Out = Val(NVL(rsTemp!异常类型)) = 2
    strTittle = IIf(Not blnDel_Out, "结帐", "重退")
    str操作员姓名 = NVL(rsTemp!操作员姓名)
    
    
    If str操作员姓名 <> UserInfo.姓名 Then
        '100703
         If MsgBox("注意:" & vbCrLf & _
                    "    该病人存在异常的" & strTittle & "单据" & IIf(str操作员姓名 <> UserInfo.姓名, ",该单据是操作员[" & str操作员姓名 & "]收取的," & vbCrLf, "") & " ,你是否继续进行" & str当前操作 & "操作?" & vbCrLf & vbCrLf & _
                    "『是』代表不对异常单据进行处理,继续进行" & str当前操作 & "操作. " & vbCrLf & _
                    "『否』代表中止结帐操作.", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            zlErrBalanceCheckFromPatiID = 1
         Else
           zlErrBalanceCheckFromPatiID = 2
         End If
        Exit Function
    End If
    If MsgBox("注意:" & vbCrLf & _
                        "       该病人存在异常的" & strTittle & "单据(" & strErrNo_Out & "),你是否需要重新对该单据进行" & strTittle & "?" & vbCrLf & vbCrLf & _
                        "『是』代表重新对异常单据进行" & strTittle & vbCrLf & _
                        "『否』代表不对异常单据进行处理,继续进行" & str当前操作 & "操作.", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        zlErrBalanceCheckFromPatiID = 1
    Else
        zlErrBalanceCheckFromPatiID = 3
    End If
End Function

Public Function zlGetBalanceIDFromBalanceNo(strNO As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐NO,获取原始的结帐ID
    '入参:strNo-结帐ID
    '返回:返回原始的结帐ID
    '编制:刘兴洪
    '日期:2020-03-25 18:52:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select ID From 病人结帐记录 Where NO=[1] And 记录状态 in (1,3)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据结帐单获取原结帐ID", strNO)
    If rsTemp.EOF Then Exit Function
     zlGetBalanceIDFromBalanceNo = Val(NVL(rsTemp!ID))
End Function


