Attribute VB_Name = "mdlInsureBalance"
Option Explicit
Public gclsInsure As New clsInsure          '医保接口对象
Public Enum 身份验证Enum
    id门诊收费 = 0
    id入院登记 = 1
    id帐户管理 = 2
    id挂号 = 3
    id结帐 = 4
    id门诊确认 = 5
End Enum

Public Enum 医院业务
    support门诊预算 = 0
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    'support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤销出院 = 17            '允许撤消病人出院
    support必须录入入出诊断 = 18    '病人入院与出院时，必须录入诊断名
    support记帐完成后上传 = 19      '要求上传在记帐数据提交后再进行
    support出院结算必须出院 = 20    '病人结帐时如果选择出院结帐，就检查必须出院才可以进行
    
    support挂号使用个人帐户 = 21    '使用医保挂号时是否使用个人帐户进行支付

    support门诊连续收费 = 22        '门诊在身份验证后，可进行多次收费操作
    support门诊收费完成后验证 = 23  '在门诊收费完成，是否再次调用身份验证
    
    support医嘱上传 = 24            '医嘱产生费用时是否实时传输
    support分币处理 = 25            '医保病人是否处理分币
    support中途结算仅处理已上传部分 = 26 '提供对已上传部分数据的结算功能
    support允许冲销已结帐的记帐单据 = 27 '是否允许冲销记帐单据，如果该单据已经结帐
    
    support允许部份冲销单据 = 28
    support出院无实际交易 = 29      '出院接口中是否要与接口商进行交易
    support多单据收费 = 30          '是否支持多单据收费
    
    support门诊收费存为划价单 = 31  '将门诊收费单转为划价单保存，修改以前固定判断某个医保的方式
    
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
    support负数记帐 = 35            '是否允许负数记帐，操作员首先要拥有负数记帐的权限。此参数缺省为真，不支持的接口需单独处理
    support多单据收费必须全退 = 39  '多单据收费必须全退
    
    support医保接口打印票据 = 46    'HIS中只走票据号但不调打印，医保接口(北京)中打印
    support多单据一次结算 = 47      '多单据预结算时，医保接口仅在最后一次调用时返回结算结果，HIS中再分摊到每张单据上
    
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    support医生确定处方类型 = 48
    support实时监控 = 60             '是否启用费用实时监控
    '刘兴洪:27536 20100119
    support不提醒缴款金额不足 = 64            '在收费时,如果收费参数的"不进行缴款输入和累计控制"为true时,同时是医保病人时没有输入缴款金额时不提醒用户
    support退费后打印回单 = 65   '医保病人是否退费后打印回单:问题
    
    support上传门诊档案 = 70                    '在门诊医嘱发送时，是否调用TranElecDossier函数完成门诊病人电子卷宗/电子档案的上传
    
    support门诊_不分单据结算 = 80               '预结算、结算都只调用一次医保交易:一卡通同步更改
    
    support挂号不收取病历费 = 81    '在挂号时，不使用医保收取病历费

    support按单据全退 = 82 '门诊退费时，按单据进行退费，86176
    support多单据分单据结算 = 83 '多单据一次结算按单据进行医保报销，86321
    support一次结算分单据退费 = 85 '按一次结算调用医保接口，但按单据退费,91602
    
    support挂号检查项目 = 86
    support门诊挂号预算 = 89
End Enum

Public Type Ty_InsurePara
    允许不设置医保项目 As Boolean
    门诊收费存为划价单 As Boolean
    不提醒缴款金额不足 As Boolean
    门诊必须传递明细 As Boolean
    医保接口打印票据 As Boolean
    
    医生确定处方类型 As Boolean
    多单据一次结算 As Boolean
    多单据分单据结算 As Boolean
    一次结算分单据退费 As Boolean
    门诊连续收费 As Boolean
    
    门诊预结算 As Boolean
    多单据收费 As Boolean
    分币处理 As Boolean
    实时监控 As Boolean
    先自付 As Boolean
    
    全自付 As Boolean
    blnOnlyBjYb As Boolean '本地仅支持北京医保:刘兴洪
    退费后打印回单 As Boolean
    医保不走票号 As Boolean
    门诊结算作废 As Boolean
    
    按单据全退 As Boolean
End Type

Public Enum 交易Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
    Busi_TransferSwap
    Busi_TransferDelSwap
    Busi_WipeoffMoney
    Busi_SettleSwap
    Busi_SettleDelSwap
    Busi_ComeInSwap
    Busi_LeaveSwap
    Busi_TranChargeDetail
    Busi_LeaveDelSwap
    Busi_RegistSwap
    Busi_RegistDelSwap
    Busi_ComeInDelSwap
    Busi_ModiPatiSwap
    Busi_ChooseDisease
    Busi_IdentifyCancel
End Enum

'结算类型
Public Enum Enum_BalanceType
    普通结算 = 0
    预存款 = 1
    医保 = 2 '不含医疗卡医保结算
    一卡通 = 3
    老一卡通 = 4
    消费卡 = 5
End Enum

Public Function initInsurePara(ByVal intInsure As Integer, ByVal lng病人ID As Long, _
    Optional ByVal lng结帐ID As Long) As Ty_InsurePara
    '初始化医保参数
    Dim tyInsurePara As Ty_InsurePara
    
    On Error GoTo ErrHandler
    If intInsure = 0 Then Exit Function
    If gclsInsure Is Nothing Then Exit Function
    
    tyInsurePara.允许不设置医保项目 = gclsInsure.GetCapability(support允许不设置医保项目, lng病人ID, intInsure)
    tyInsurePara.门诊收费存为划价单 = gclsInsure.GetCapability(support门诊收费存为划价单, lng病人ID, intInsure)
    tyInsurePara.门诊必须传递明细 = gclsInsure.GetCapability(support门诊必须传递明细, lng病人ID, intInsure)
    tyInsurePara.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
    
    tyInsurePara.多单据一次结算 = gclsInsure.GetCapability(support多单据一次结算, lng病人ID, intInsure)
    tyInsurePara.多单据分单据结算 = gclsInsure.GetCapability(support多单据分单据结算, lng病人ID, intInsure)
    tyInsurePara.一次结算分单据退费 = gclsInsure.GetCapability(support一次结算分单据退费, lng病人ID, intInsure)
    tyInsurePara.门诊连续收费 = gclsInsure.GetCapability(support门诊连续收费, lng病人ID, intInsure)
    
    tyInsurePara.门诊预结算 = gclsInsure.GetCapability(support门诊预算, lng病人ID, intInsure)
    tyInsurePara.多单据收费 = gclsInsure.GetCapability(support多单据收费, lng病人ID, intInsure)
    tyInsurePara.分币处理 = gclsInsure.GetCapability(support分币处理, lng病人ID, intInsure)
    tyInsurePara.实时监控 = gclsInsure.GetCapability(support实时监控, lng病人ID, intInsure)
    tyInsurePara.先自付 = gclsInsure.GetCapability(support收费帐户首先自付, lng病人ID, intInsure)
    
    tyInsurePara.全自付 = gclsInsure.GetCapability(support收费帐户全自费, lng病人ID, intInsure)
    tyInsurePara.blnOnlyBjYb = False
    tyInsurePara.医保不走票号 = False
    initInsurePara = tyInsurePara
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlExecuteInsurePreSwap(ByVal bytMode As Byte, objBalanceBills As BalanceBills, _
    ByVal intInsure As Integer, ByRef colBalance As Collection, _
    ByVal strErrMsg As String, _
    Optional ByVal blnErrBill As Boolean) As Boolean
    '门诊预结算
    '入参：
    '   bytMode 医保结算模式：0-多单据一次结算,1-多单据一次结算分单据退费,2-多单据分单据结算
    '   objBalanceBills 费用数据
    '   strInvoice 当前发票号
    '出参：
    '   colBalance 预结算结果集(每张单据对应一个BalanceMoneys对象元素),多单据一次结算时存在第一张单据中
    '   strErrMsg 错误信息,False时返回
    Dim strDate As String, rsBalance As ADODB.Recordset
    Dim rsRecord As ADODB.Recordset
    Dim strBalance As String, strAdvance As String
    Dim varBalance As Variant, varItem As Variant, str结算方式 As String
    Dim p As Long, i As Long
    Dim strNos As String
    
    On Error GoTo ErrHandler
    strErrMsg = ""
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:mm:ss")
    
    For p = 1 To objBalanceBills.Count
        strNos = strNos & "," & objBalanceBills(p).NO
    Next
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    '2-多单据分单据结算
    If bytMode = 2 Then
        If blnErrBill Then Set rsBalance = zlGetBalanceDetail(2, Mid(strNos, 2), 1)
        For p = 1 To objBalanceBills.Count
            strBalance = ""
            If blnErrBill Then
                '检查该张单据是否已成功医保结算
                rsBalance.Filter = "No='" & objBalanceBills(p).NO & "'"
                Do While Not rsBalance.EOF
                    strBalance = strBalance & IIf(strBalance = "", "", "||")
                    strBalance = strBalance & Nvl(rsBalance!结算方式) & "|" & Val(Nvl(rsBalance!金额))
                    rsBalance.MoveNext
                Loop
            End If
            
            If strBalance <> "" Then
                Call SetBalanceVal(colBalance, p, strBalance)
            Else
                Set rsRecord = MakePreSwapDataFromDB(objBalanceBills(p).NO)
                
                strBalance = "": strAdvance = ""
                If Not gclsInsure.ClinicPreSwap(rsRecord, strBalance, intInsure, strAdvance) Then
                    strErrMsg = "第 " & p & " 张单据预结算失败。"
                    Exit Function
                End If
                
                '只要有一张单据自动走票号，都要走票号
                'If strAdvance <> "" And InStr(strAdvance, "|") = 0 Then    '医保票据号 Then
                '    '38821,格式:票据号;是否不走票号(1-不走票号;0-自动走票号)
                '    varItem = Split(strAdvance & ";", ";")
                '    strInsureInvoice = varItem(0)
                '    bln不走票号 = bln不走票号 And Val(varItem(1)) = 1
                'End If
                
                '报销方式;金额;是否允许修改|....
                If strBalance <> "" Then
                    strBalance = Replace(Replace(strBalance, "|", "||"), ";", "|")
                    Call SetBalanceVal(colBalance, p, strBalance)
                End If
            End If
        Next
        ZlExecuteInsurePreSwap = True: Exit Function
    End If
    
    
    '0-多单据一次结算,1-多单据一次结算分单据退费
    Set rsRecord = MakePreSwapDataFromDB(strNos)
    
    strBalance = "": strAdvance = ""
    If Not gclsInsure.ClinicPreSwap(rsRecord, strBalance, intInsure, strAdvance) Then
        strErrMsg = "单据预结算失败。"
        Exit Function
    End If
    
    'If strAdvance <> "" And InStr(strAdvance, "|") = 0 Then
    '    '38821:strAdvance:发票号;是否不走票据号
    '    varItem = Split(strAdvance & ";", ";")
    '    strInsureInvoice = varItem(0)
    '    bln不走票号 = Val(varItem(1)) = 1
    'End If
    
    '报销方式;金额;是否允许修改|....
    If strBalance <> "" Then
        If bytMode = 0 Then
            '0-多单据一次结算
            strBalance = Replace(Replace(strBalance, "|", "||"), ";", "|")
            Call SetBalanceVal(colBalance, 1, strBalance)
        Else
            '1-多单据一次结算分单据退费
            '单据序号:结算方式;金额;是否允许修改|...||单据序号:结算方式;金额;是否允许修改|...||...
            varBalance = Split(strBalance, "||")
            For i = 0 To UBound(varBalance)
                If InStr(varBalance(i), ":") = 0 Then
                    strErrMsg = "单据预结算返回结算结果格式不正确。"
                    Exit Function
                End If
                
                varItem = Split(varBalance(i), ":")
                p = Val(varItem(0)): str结算方式 = varItem(1)
                If p < 1 Or p > colBalance.Count Then
                    strErrMsg = "单据预结算返回结算结果格式不正确。"
                    Exit Function
                End If
                
                str结算方式 = Replace(Replace(str结算方式, "|", "||"), ";", "|")
                Call SetBalanceVal(colBalance, p, str结算方式)
            Next
        End If
    End If
    
    ZlExecuteInsurePreSwap = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckInsureBalanceValid(ByRef rs结算方式 As ADODB.Recordset, _
    ByVal colBalance As Collection) As String
    '检查医保有但本地没有的结算方式，返回本地没有的结算方式
    '入参：
    '   colBalance BalanceMoneys对象
    Dim i As Integer, strNone As String
    Dim objItem As BalanceMoney
    
    On Error GoTo ErrHandler
    If colBalance Is Nothing Then Exit Function
    
    For i = 1 To colBalance.Count
        For Each objItem In colBalance(i)
            If rs结算方式 Is Nothing Then
                If InStr("," & strNone & ",", "," & objItem.结算方式 & ",") = 0 Then
                    strNone = strNone & "," & objItem.结算方式
                End If
            Else
                rs结算方式.Filter = "(名称='" & objItem.结算方式 & "' And 性质=3) Or (名称='" & objItem.结算方式 & "' And 性质=4)"
                If rs结算方式.EOF Then
                    If InStr("," & strNone & ",", "," & objItem.结算方式 & ",") = 0 Then
                        strNone = strNone & "," & objItem.结算方式
                    End If
                End If
            End If
        Next
    Next
    If strNone <> "" Then strNone = Mid(strNone, 2)
    
    CheckInsureBalanceValid = strNone
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InsureBalanced(ByVal intInsure As Integer, ByVal lng结帐ID As Long) As Boolean
    '判断是否已成功进行了医保结算
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If intInsure = 0 Then Exit Function
    '校对标志等于2则已成功结算
    strSql = _
        "Select 1" & vbNewLine & _
        "From 病人预交记录 A, 结算方式 B" & vbNewLine & _
        "Where a.结算方式 = b.名称 And b.性质 In (3, 4)  And Nvl(校对标志, 0) = 2" & vbNewLine & _
        "      And a.记录性质 = 3 And a.结帐id = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "mdlCliniBalance", lng结帐ID)
    InsureBalanced = Not rsTemp.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetMedicareSum(colBalance As Collection, Optional ByVal strItem As String, Optional ByVal intPage As Integer, _
    Optional ByVal blnOrig As Boolean, Optional ByVal intPageCount As Integer) As Currency
    '功能：获取保险结算的金额
    '参数：strItem=是否指定结算方式,否则为所有结算方式
    '      blnOrig=是否取原始(最大)结算金额,否则取现在(修改后)有效金额
    '      intPage=是否指定单据,否则为所有单据
    '      intPageCount=结算单据张数
    '说明：该函数以colBalance为准计算,对于医保划价收费也是
    Dim curMoney As Currency, p As Integer
    Dim intPageStart As Integer, intPageEnd As Integer
    Dim objItem As BalanceMoney
    
    intPageStart = IIf(intPage = 0, 1, intPage)
    intPageEnd = IIf(intPage = 0, IIf(intPageCount = 0, colBalance.Count, intPageCount), intPage)
    For p = intPageStart To intPageEnd
        For Each objItem In colBalance(p)
            If strItem = "" Or objItem.结算方式 = strItem Then
                If blnOrig Then
                    curMoney = curMoney + objItem.原始金额
                Else
                    curMoney = curMoney + objItem.有效金额
                End If
            End If
        Next
    Next
    GetMedicareSum = curMoney
End Function

Public Function GetMedicareStr(colBalance As Collection, Optional ByVal intPage As Integer, _
    Optional ByVal intPageCount As Integer) As String
    '功能：返回保险结算方式串,"结算方式|金额||...."
    '参数：intPage=是否指定单据,否则为所有单据
    '      intPageCount=结算单据总张数
    '说明：该函数以colBalance为准计算,对于医保划价收费也是
    Dim p As Integer
    Dim rsTemp As New ADODB.Recordset, strBalance As String
    Dim intPageStart As Integer, intPageEnd As Integer
    Dim objItem As BalanceMoney
    
    On Error GoTo ErrHander
    rsTemp.Fields.Append "结算方式", adVarChar, 20, adFldIsNullable
    rsTemp.Fields.Append "金额", adCurrency, , adFldIsNullable
    rsTemp.CursorLocation = adUseClient
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorType = adOpenStatic
    rsTemp.Open
    
    intPageStart = IIf(intPage = 0, 1, intPage)
    intPageEnd = IIf(intPage = 0, IIf(intPageCount = 0, colBalance.Count, intPageCount), intPage)
    For p = intPageStart To intPageEnd
        For Each objItem In colBalance(p)
            rsTemp.Find "结算方式='" & objItem.结算方式 & "'", , adSearchForward, 1
            If rsTemp.EOF Then rsTemp.AddNew
            rsTemp!结算方式 = objItem.结算方式
            rsTemp!金额 = Val(Nvl(rsTemp!金额)) + objItem.有效金额
            rsTemp.Update
        Next
    Next
    
    strBalance = ""
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        strBalance = strBalance & "||" & Nvl(rsTemp!结算方式) & "|" & Nvl(rsTemp!金额)
        rsTemp.MoveNext
    Loop
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    
    GetMedicareStr = strBalance
    Exit Function
ErrHander:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetInsureBalanceSum(objBalanceMoneys As BalanceMoneys, _
    Optional ByVal strItem As String, Optional ByVal blnOrig As Boolean) As Currency
    '获取保险结算的金额
    '入参：
    '   strItem 是否指定结算方式,否则为所有结算方式
    '   blnOrig 是否取原始(最大)结算金额,否则取现在(修改后)有效金额
    Dim curMoney As Currency
    Dim objItem As BalanceMoney
    
    On Error GoTo ErrHander
    If objBalanceMoneys Is Nothing Then Exit Function
    For Each objItem In objBalanceMoneys
        If strItem = "" Or objItem.结算方式 = strItem Then
            If blnOrig Then
                curMoney = curMoney + objItem.原始金额
            Else
                curMoney = curMoney + objItem.有效金额
            End If
        End If
    Next
    GetInsureBalanceSum = curMoney
    Exit Function
ErrHander:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetInsureBalanceStrAll(objBalanceBills As BalanceBills) As String
    '获取所有单据的预结算结果,"结算方式|金额||...."
    Dim i As Integer
    Dim colBalance As New Collection
    
    If objBalanceBills Is Nothing Then Exit Function
    For i = 1 To objBalanceBills.Count
        colBalance.Add objBalanceBills(i).预结算
    Next
    GetInsureBalanceStrAll = GetMedicareStr(colBalance)
End Function

Public Function GetInsureBalanceStr(objBalanceMoneys As BalanceMoneys) As String
    '获取保险结算串,"结算方式|金额||...."
    Dim strBalance As String
    Dim objItem As BalanceMoney
    
    On Error GoTo ErrHander
    If objBalanceMoneys Is Nothing Then Exit Function
    For Each objItem In objBalanceMoneys
        strBalance = strBalance & "||" & objItem.结算方式 & "|" & objItem.有效金额
    Next
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    
    GetInsureBalanceStr = strBalance
    Exit Function
ErrHander:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetBalanceVal(colBalance As Collection, ByVal intPage As Integer, _
    ByVal strBalance As String)
    '功能：设置指定编号单据指定保险结算方式的有效值
    '参数：
    '       strBalance-根据结算方式字符串设置结算方式记录集，格式：结算方式1|金额1||结算方式2|金额2||...
    '说明：该函数以colBalance为准计算,对于医保划价收费也是
    '说明：用于正常医保收费修改保险结算金额；及划价单医保收费设置个人帐户等结算金额
    Dim i As Long
    Dim varBalance As Variant, varTemp As Variant
    Dim blnFind As Boolean
    Dim objItem As BalanceMoney, objBalanceMoneys As BalanceMoneys
    
    If strBalance = "" Then Exit Sub
    
    Set objBalanceMoneys = colBalance(intPage)
    
    '格式：结算方式1|金额1||结算方式2|金额2||...
    varBalance = Split(strBalance, "||")
    For i = 0 To UBound(varBalance)
        varTemp = Split(varBalance(i) & "|||", "|")
        blnFind = False
        For Each objItem In objBalanceMoneys
            If objItem.结算方式 = varTemp(0) Then
                objItem.有效金额 = varTemp(1)
                blnFind = True: Exit For
            End If
        Next
            
        If Not blnFind Then
            Set objItem = New BalanceMoney
            objItem.结算方式 = varTemp(0)
            objItem.原始金额 = varTemp(1)
            objItem.允许修改 = Val(varTemp(2)) = 1
            objItem.有效金额 = varTemp(1)
            objBalanceMoneys.AddItem objItem
        End If
    Next

    colBalance.Remove intPage '集合元素不能直接修改
    If colBalance.Count >= intPage Then
        colBalance.Add objBalanceMoneys, , intPage
    Else
        colBalance.Add objBalanceMoneys
    End If
End Sub

Public Function zlInsureCheck(ByVal str预结算 As String, ByVal strAdvance As String) As Boolean
    '检查当前的医保是否需要较对
    '入参:
    '   str预结算-保险结算
    '   strAdvance-医保返回的结算
    '说明：
    '   正式结算前后,结算方式和结算金额未发生变化时不校对
    Dim blnFind  As Boolean, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo ErrHandler
    If strAdvance = "" Or str预结算 = strAdvance Then Exit Function
    
    zlInsureCheck = True
    
    varData = Split(str预结算, "||")
    varData1 = Split(strAdvance, "||")
    If UBound(varData) <> UBound(varData1) Then Exit Function
    
    For i = 0 To UBound(varData)
        blnFind = False
        varTemp = Split(varData(i), "|")
        For j = 0 To UBound(varData1)
            varTemp1 = Split(varData1(j), "|")
            If varTemp(0) = varTemp1(0) Then
                blnFind = True
                If Val(varTemp(1)) <> Val(varTemp1(1)) Then Exit Function
            End If
        Next
        If Not blnFind Then Exit Function
    Next
    zlInsureCheck = False
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlExecuteInsureSwap(ByVal bytMode As Byte, ByVal objPati As clsPatientInfo, _
    ByVal intInsure As Integer, ByVal str个账名称 As String, ByVal blnOnlyBalanceSuccessedNo As Boolean, _
    ByVal lng结帐ID As Long, ByVal lng结算序号 As Long, objBalanceBills As BalanceBills, _
    ByRef blnCommit As Boolean, Optional ByRef strSavedNos As String, Optional ByRef lngSavedBillCount As Long, _
    Optional ByRef blnYbBalanced As Boolean, Optional ByRef strErrMsg As String) As Boolean
    '医保结算
    '入参：
    '   bytMode 医保结算模式：0-多单据一次结算,1-多单据一次结算分单据退费,2-多单据分单据结算
    '   blnOnlyBalanceSuccessedNo 多单据分单据结算时是否只对医保结算成功单据收费
    '   strSavedNos,lngSavedBillCount 多单据分单据结算时已结算成功的单据情况
    '   blnYbBalanced 多单据分单据结算时对医保结算成功单据收费
    '说明:需要在外层启用事务,正常退费后,该过程已提交,不需要调用者提交
    '     如果失败,则事务将回退(主要是避免弹出界面造成死锁)
    Dim colBalance As Collection, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim strAdvance As String, strAdvanceOld  As String
    Dim cur个帐支付 As Currency, cur医保基金 As Currency
    Dim cur全自付 As Currency, cur先自付 As Currency
    Dim strAll预结算 As String, str预结算 As String, str结算方式 As String
    Dim rsBalance As ADODB.Recordset, objBill As BalanceBill
    Dim p As Long, i As Long, blnFind As Boolean
    Dim varAdvance As Variant, varItem As Variant
    Dim blnCurrentCommit As Boolean
    
    On Error GoTo ErrHandler
    blnCommit = False: strSavedNos = ""
    blnYbBalanced = False: strErrMsg = ""
    If intInsure = 0 Then gcnOracle.RollbackTrans: Exit Function
    
    blnTrans = True
    strAll预结算 = GetInsureBalanceStrAll(objBalanceBills)
    '先保存预结算结果
    Call SaveInsureBalance(objPati, lng结帐ID, strAll预结算)
    
    '2-多单据分单据结算
    If bytMode = 2 Then
        Set colBalance = New Collection
        Set rsBalance = zlGetBalanceDetail(0, lng结帐ID, 1)
        
        For p = 1 To objBalanceBills.Count
            colBalance.Add New BalanceMoneys
            Set objBill = objBalanceBills(p)
            
            '检查该张单据是否已成功医保结算
            str结算方式 = GetYBBalanceNo(rsBalance, objBill.NO)
            
            If str结算方式 <> "" Then
                Call SetBalanceVal(colBalance, p, str结算方式)
                strSavedNos = strSavedNos & "," & objBill.NO
            Else
                strAdvance = lng结算序号 & "|" & objBill.NO
                strAdvanceOld = strAdvance
                
                str预结算 = GetInsureBalanceStr(objBill.预结算)
                Call SaveInsureBalanceDetail(lng结帐ID, objBill.NO, str预结算)
                
                cur个帐支付 = GetInsureBalanceSum(objBill.预结算, str个账名称)
                cur医保基金 = GetInsureBalanceSum(objBill.预结算, "医保基金")
                cur全自付 = objBill.全自付
                cur先自付 = objBill.先自付
                
                If Not gclsInsure.ClinicSwap(lng结帐ID, cur个帐支付, cur医保基金, cur全自付, cur先自付, _
                    intInsure, strAdvance) Then
                    If blnOnlyBalanceSuccessedNo Then GoTo ErrHandler:
                    gcnOracle.RollbackTrans
                    If blnCurrentCommit Then Call CorrectInsureErrBalance(objPati, lng结帐ID)  '医保结算校对
                    Exit Function
                End If
                If strAdvance = strAdvanceOld Then strAdvance = ""
                blnTransMedicare = True
                
                If zlInsureCheck(str预结算, strAdvance) Then
                    Call SaveInsureBalanceDetail(lng结帐ID, objBill.NO, strAdvance)
                    str预结算 = strAdvance
                End If
                
                Call SetBalanceVal(colBalance, p, str预结算)
                strSavedNos = strSavedNos & "," & objBill.NO
                
                gcnOracle.CommitTrans: blnTrans = False
                blnCommit = True: blnCurrentCommit = True
                
                Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, intInsure)
                blnTransMedicare = False
                
                gcnOracle.BeginTrans: blnTrans = True
            End If
        Next
        strAdvance = GetMedicareStr(colBalance)
        
    '1-多单据一次结算分单据退费
    ElseIf bytMode = 1 Then
        Set colBalance = New Collection
        strAdvance = lng结算序号
        strAdvanceOld = strAdvance
        
        For p = 1 To objBalanceBills.Count
            Set objBill = objBalanceBills(p)
            str预结算 = GetInsureBalanceStr(objBill.预结算)
            Call SaveInsureBalanceDetail(lng结帐ID, objBill.NO, str预结算)
            
            cur个帐支付 = cur个帐支付 + GetInsureBalanceSum(objBill.预结算, str个账名称)
            cur医保基金 = cur医保基金 + GetInsureBalanceSum(objBill.预结算, "医保基金")
            cur全自付 = cur全自付 + objBill.全自付
            cur先自付 = cur先自付 + objBill.先自付
        Next
        
        If Not gclsInsure.ClinicSwap(lng结帐ID, cur个帐支付, cur医保基金, cur全自付, cur先自付, _
            intInsure, strAdvance) Then gcnOracle.RollbackTrans: Exit Function
        If strAdvance = strAdvanceOld Then strAdvance = ""
        blnTransMedicare = True
        
        'NO:结算方式,金额|结算方式,金额|...||NO:结算方式,金额|结算方式,金额|...||...
        varAdvance = Split(strAdvance, "||")
        For p = 1 To objBalanceBills.Count
            Set objBill = objBalanceBills(p)
            '如果其中某一张单据没有返回对应结算信息，就按预结算结果保存
            blnFind = False
            For i = 0 To UBound(varAdvance)
                If InStr(varAdvance(i), ":") = 0 Then
                    strErrMsg = "医保结算结果格式不正确！"
                    Exit Function
                End If
                
                varItem = Split(varAdvance(i), ":")
                If objBill.NO = varItem(0) Then
                    str结算方式 = Replace(Replace(varItem(1), "|", "||"), ",", "|")
                    blnFind = True: Exit For
                End If
            Next
            
            If blnFind Then
                '直接修正医保结果，不检查是否需要校对
                Call SaveInsureBalanceDetail(lng结帐ID, objBill.NO, str结算方式)
            Else
                str结算方式 = GetInsureBalanceStr(objBill.预结算)
            End If
            
            colBalance.Add New BalanceMoneys
            SetBalanceVal colBalance, p, str结算方式
        Next
        strAdvance = GetMedicareStr(colBalance)
    
    '0-多单据一次结算
    Else
        strAdvance = lng结算序号
        strAdvanceOld = strAdvance
        
        For p = 1 To objBalanceBills.Count
            Set objBill = objBalanceBills(p)
            cur个帐支付 = cur个帐支付 + GetInsureBalanceSum(objBill.预结算, str个账名称)
            cur医保基金 = cur医保基金 + GetInsureBalanceSum(objBill.预结算, "医保基金")
            cur全自付 = cur全自付 + objBill.全自付
            cur先自付 = cur先自付 + objBill.先自付
        Next
        
        If Not gclsInsure.ClinicSwap(lng结帐ID, cur个帐支付, cur医保基金, _
            cur全自付, cur先自付, intInsure, strAdvance) Then gcnOracle.RollbackTrans: Exit Function
        If strAdvance = strAdvanceOld Then strAdvance = ""
        blnTransMedicare = True
    End If
    
    '校对整体的结算结果
    If zlInsureCheck(strAll预结算, strAdvance) Then
        Call SaveInsureBalance(objPati, lng结帐ID, strAdvance)
    End If
    Call InsureBalanceOver(lng结帐ID)
    gcnOracle.CommitTrans: blnTrans = False
    
    If blnTransMedicare Then
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, intInsure)
    End If
    zlExecuteInsureSwap = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
    If blnTransMedicare Then
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, False, intInsure)
    End If
    
    If bytMode = 2 And strSavedNos <> "" Then
        '105338:部分结算成功，只对结算成功这部分单据收费
        If blnOnlyBalanceSuccessedNo Then
            On Error GoTo LastErrHandler
            strSavedNos = Mid(strSavedNos, 2)
            lngSavedBillCount = p - 1
            
            strAdvance = GetMedicareStr(colBalance)
            gcnOracle.BeginTrans: blnTrans = True
            '1.删除未成功的费用单据，恢复为划价单
            For i = objBalanceBills.Count To p Step -1
                Set objBill = objBalanceBills(i)
                If InStr("," & strSavedNos & ",", "," & objBill.NO & ",") = 0 Then
                    Call CancelBillBalance(lng结帐ID, objBill.NO)
                End If
            Next
            
            '2.校对医保结算
            Call SaveInsureBalance(objPati, lng结帐ID, strAdvance)
            Call InsureBalanceOver(lng结帐ID)
            gcnOracle.CommitTrans: blnTrans = False
            blnYbBalanced = True: Exit Function
        ElseIf blnCurrentCommit Then
            Call CorrectInsureErrBalance(objPati, lng结帐ID) '医保结算校对
        End If
    ElseIf Err <> 0 Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    Exit Function
LastErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SaveInsureBalanceDetail(ByVal lng结帐ID As Long, ByVal strNO As String, _
    ByVal strBalance As String, Optional cllPro As Collection)
    '保存医保结算明细
    Dim strSql As String
    On Error GoTo errH
    
    'Zl_医保结算明细_Insert(
    strSql = "Zl_医保结算明细_Insert( "
    '  结帐id_In   医保结算明细.结帐id%Type,
    strSql = strSql & "" & lng结帐ID & ","
    '  No_In       医保结算明细.No%Type,
    strSql = strSql & "'" & strNO & "',"
    '  结算方式_In Varchar2,
    strSql = strSql & "'" & strBalance & "',"
    '  备注_In     医保结算明细.备注%Type := Null,
    strSql = strSql & "" & "NULL" & ")"
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "mdlInsureBalance"
    Else
        zlAddArray cllPro, strSql
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub SaveInsureBalance(ByVal objPati As clsPatientInfo, ByVal lng结帐ID As Long, _
    ByVal strBalance As String, Optional ByVal blnDel As Boolean, _
    Optional ByVal lng关联交易ID As Long, Optional cllPro As Collection)
    '保存医保结算数据
    Dim strSql As String
    On Error GoTo errH
    
    If blnDel Then
        'Zl_门诊退费结算_Modify_S(
        strSql = "Zl_门诊退费结算_Modify_S("
        '  操作类型_In      Number,
        strSql = strSql & "" & 3 & ","
        '  病人id_In        门诊费用记录.病人id%Type,
        strSql = strSql & "" & objPati.病人ID & ","
        '  姓名_In          病人预交记录.姓名%Type,
        strSql = strSql & "'" & objPati.姓名 & "',"
        '  性别_In          病人预交记录.性别%Type,
        strSql = strSql & "'" & objPati.性别 & "',"
        '  年龄_In          病人预交记录.年龄%Type,
        strSql = strSql & "'" & objPati.年龄 & "',"
        '  门诊号_In        病人预交记录.门诊号%Type,
        strSql = strSql & "'" & objPati.门诊号 & "',"
        '  住院号_In        病人预交记录.住院号%Type,
        strSql = strSql & "'" & objPati.住院号 & "',"
        '  付款方式名称_In  病人预交记录.付款方式名称%Type,
        strSql = strSql & "'" & objPati.医疗付款方式 & "',"
        '  冲销id_In        病人预交记录.结帐id%Type,
        strSql = strSql & "" & lng结帐ID & ","
        '  结算方式_In      Varchar2
        strSql = strSql & "'" & strBalance & "',"
        '  冲预交_In        病人预交记录.冲预交%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  卡类别id_In      病人预交记录.卡类别id%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  卡号_In          病人预交记录.卡号%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  交易流水号_In    病人预交记录.交易流水号%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  交易说明_In      病人预交记录.交易说明%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  缴款_In          病人预交记录.缴款%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  找补_In          病人预交记录.找补%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  误差金额_In      门诊费用记录.实收金额%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  完成退费_In      Number := 0,
        strSql = strSql & "" & "0" & ","
        '  原结帐id_In      病人预交记录.结帐id%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  剩余转预交_In    Number := 0,
        strSql = strSql & "" & "0" & ","
        '  缺省结算方式_In  结算方式.名称%Type := Null,
        strSql = strSql & "" & "NULL" & ","
        '  冲预交病人ids_In Varchar2 := Null,
        strSql = strSql & "" & "NULL" & ","
        '  关联交易id_In    病人预交记录.关联交易id%Type := Null,
        strSql = strSql & "" & IIf(lng关联交易ID = 0, "NULL", lng关联交易ID) & ")"
    Else
        'Zl_门诊收费结算_Modify_S
        strSql = "Zl_门诊收费结算_Modify_S("
        '  操作类型_In   Number,
        strSql = strSql & "" & 2 & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSql = strSql & "" & objPati.病人ID & ","
        '  姓名_In          病人预交记录.姓名%Type,
        strSql = strSql & "'" & objPati.姓名 & "',"
        '  性别_In          病人预交记录.性别%Type,
        strSql = strSql & "'" & objPati.性别 & "',"
        '  年龄_In          病人预交记录.年龄%Type,
        strSql = strSql & "'" & objPati.年龄 & "',"
        '  门诊号_In        病人预交记录.门诊号%Type,
        strSql = strSql & "'" & objPati.门诊号 & "',"
        '  住院号_In        病人预交记录.住院号%Type,
        strSql = strSql & "'" & objPati.住院号 & "',"
        '  付款方式名称_In  病人预交记录.付款方式名称%Type,
        strSql = strSql & "'" & objPati.医疗付款方式 & "',"
        '  结帐id_In     病人预交记录.结帐id%Type,
        strSql = strSql & "" & lng结帐ID & ","
        '  结算方式_In   Varchar2,
        strSql = strSql & "'" & strBalance & "')"
    End If
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "mdlInsureBalance"
    Else
        zlAddArray cllPro, strSql
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub InsureBalanceOver(ByVal lng结帐ID As Long, _
    Optional cllPro As Collection)
    '医保完成结算，更新校对标志
    Dim strSql As String
    On Error GoTo errH
    
    'Zl_病人门诊收费_医保更新(
    strSql = "Zl_病人门诊收费_医保更新( "
    '  结帐id_In   门诊费用记录.结帐id%Type,
    strSql = strSql & "" & lng结帐ID & ","
    '  结算序号_In 病人预交记录.结算序号%Type,
    strSql = strSql & "" & "NULL" & ","
    '  保险结算_In Varchar2
    strSql = strSql & "" & "NULL" & ")"
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "mdlInsureBalance"
    Else
        zlAddArray cllPro, strSql
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function Get个帐报销金额(ByVal cur实收合计 As Currency, ByVal cur个帐预付 As Currency, _
    ByVal dbl个帐余额 As Double, ByVal dbl个帐透支 As Double) As Currency
    '计算个人帐户支付金额
    If RoundEx(cur实收合计, 6) <= 0 Then Get个帐报销金额 = 0: Exit Function
    
    If RoundEx(dbl个帐余额 + dbl个帐透支, 6) <= 0 Then '当前已无余额(含透支)
        Get个帐报销金额 = 0
    Else
        If RoundEx(dbl个帐余额 + dbl个帐透支, 6) >= RoundEx(cur个帐预付, 6) Then '在允许支付范围内足够(含透支)
            Get个帐报销金额 = cur个帐预付
        Else
            Get个帐报销金额 = dbl个帐余额 + dbl个帐透支
        End If
    End If
End Function

Public Function CorrectInsureErrBalance(ByVal objPati As clsPatientInfo, _
    ByVal lng结帐ID As Long, Optional ByVal blnDel As Boolean) As Boolean
    '医保结算校对
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim rsBalance As ADODB.Recordset, strBalance As String
    Dim rsBalanceSaved As ADODB.Recordset, strBalanceSaved As String
    
    On Error GoTo ErrHandler
    strSql = "Select 1" & _
            " From 病人预交记录 A, 结算方式 B" & _
            " Where a.结算方式 = b.名称 And b.性质 In (3, 4) And 结帐id = [1] And a.卡类别ID Is Null " & _
            "       And Nvl(a.校对标志, 0) = 1 And Rownum < 2"
    strSql = strSql & "Union All" & _
            " Select 1" & _
            " From 保险结算记录" & _
            " Where 记录id = [1] " & _
            "       And Not Exists(Select 1 " & _
            "                      From 病人预交记录 A, 结算方式 B" & _
            "                      Where a.结算方式 = b.名称 And a.结帐id = 记录id " & _
            "                            And b.性质 In (3, 4) And a.卡类别ID Is Null)" & _
            "       And 卡类别ID Is Null And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "检查是否存在需要校对的医保结算", lng结帐ID)
    If rsTemp.EOF Then CorrectInsureErrBalance = True: Exit Function
    
    '先通过“医保结算明细”进行校对
    Set rsBalance = zlGetBalanceDetail(0, lng结帐ID, 1)
    strBalance = GetYBBalanceNo(rsBalance)
    
    If strBalance = "" Then
        strSql = "Select a.结帐ID,a.结算方式,a.金额" & _
            " From 保险结算明细 A ,结算方式 C" & _
            " Where a.结算方式=c.名称 And c.性质 in (3,4) And a.结帐id =[1] And a.标志=1 " & _
            " Order by 结算方式"
        '医保管控的过程固定写入了一条"现金",所以排开非医保类的结算方式
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", lng结帐ID)
        Do While Not rsTemp.EOF
            strBalance = strBalance & "||" & Nvl(rsTemp!结算方式) & "|" & Val(Nvl(rsTemp!金额))
            rsTemp.MoveNext
        Loop
        If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    End If
    '没有核对数据,直接返回
    If strBalance = "" Then CorrectInsureErrBalance = True: Exit Function
    
    '检查是否需要校对
    Set rsBalanceSaved = GetChargeBalance(lng结帐ID)
    strBalanceSaved = GetYBBalance(rsBalanceSaved, lng结帐ID)
    If zlInsureCheck(strBalanceSaved, strBalance) Then
        Call SaveInsureBalance(objPati, lng结帐ID, strBalance, blnDel)
    End If
    
    CorrectInsureErrBalance = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function MakePreSwapData() As ADODB.Recordset
    '创建一个预结算记录结构
    '返回:医保相关数据的数据集结构
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandler
    With rsTmp.Fields
        .Append "单据序号", adBigInt, 50, adFldIsNullable
        .Append "费别", adVarChar, 50, adFldIsNullable
        .Append "NO", adVarChar, 8, adFldIsNullable
        .Append "序号", adBigInt, , adFldIsNullable '问题:42961
        .Append "实际票号", adVarChar, 20, adFldIsNullable
        .Append "结算时间", adDBTimeStamp, , adFldIsNullable
        .Append "病人ID", adBigInt, , adFldIsNullable
        .Append "收费类别", adVarChar, 2, adFldIsNullable
        .Append "收据费目", adVarChar, 20, adFldIsNullable
        .Append "计算单位", adVarChar, 50, adFldIsNullable
        .Append "开单人", adVarChar, 100, adFldIsNullable
        .Append "收费细目ID", adBigInt, , adFldIsNullable
        .Append "数量", adDouble, , adFldIsNullable
        .Append "单价", adDouble, , adFldIsNullable
        .Append "实收金额", adCurrency, , adFldIsNullable
        .Append "统筹金额", adCurrency, , adFldIsNullable
        .Append "保险支付大类ID", adBigInt, , adFldIsNullable
        .Append "是否医保", adBigInt, , adFldIsNullable
        .Append "保险编码", adVarChar, 50, adFldIsNullable
        .Append "摘要", adVarChar, 2000, adFldIsNullable
        .Append "是否急诊", adBigInt, , adFldIsNullable
        .Append "开单部门ID", adBigInt, , adFldIsNullable
        .Append "执行部门ID", adBigInt, , adFldIsNullable
    End With
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set MakePreSwapData = rsTmp
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MakePreSwapDataFromDB(ByVal strNos As String) As ADODB.Recordset
    '根据单据对象内容创建一个记录信息(以售价单位)，主要针对全退重结和补结算
    '入参:
    '   strNos 费用单据，格式：A001,A002,...
    '出参:
    '返回:医保相关数据的数据集(单据序号(1--n),病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保)
    Dim p As Integer, strSql As String
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    
    Err = 0: On Error GoTo errHand:
    Set rsTmp = MakePreSwapData()

    strSql = _
        "Select /*+cardinality(b,10)*/a.No, Nvl(a.价格父号, a.序号) As 序号, To_Char(a.登记时间, 'YYYY-MM-DD HH24:MI:SS') As 结算时间," & vbNewLine & _
        "       a.病人id, a.费别, a.收费类别, a.收据费目, a.计算单位, a.开单人, a.收费细目id, a.保险大类id As 保险支付大类id," & vbNewLine & _
        "       Nvl(a.保险项目否, 0) As 是否医保, a.保险编码, Nvl(a.付数, 0) * a.数次 As 数量, a.标准单价 As 单价," & vbNewLine & _
        "       a.实收金额, a.统筹金额, a.摘要 As 摘要,Nvl(a.加班标志, 0) As 是否急诊, a.开单部门id, a.执行部门id, a.结帐id" & vbNewLine & _
        "From 门诊费用记录 A,(Select Column_Value As No From Table(f_Str2list([1]))) B" & vbNewLine & _
        "Where a.No = b.No And a.记录性质 = 1"
    
    strSql = _
        "Select '' As 实际票号, a.No, a.序号, Max(a.结算时间) As 结算时间, a.病人id, a.费别, a.收费类别, a.收据费目," & vbNewLine & _
        "       a.计算单位, a.开单人, a.收费细目id, a.保险支付大类id, a.是否医保, a.保险编码, Sum(a.数量) As 数量," & vbNewLine & _
        "       Max(a.单价) As 单价, Sum(a.实收金额) As 实收金额, Sum(a.统筹金额) As 统筹金额, Max(a.摘要) As 摘要," & vbNewLine & _
        "       Max(a.是否急诊) As 是否急诊, Max(a.开单部门id) As 开单部门id, Max(a.执行部门id) As 执行部门id" & vbNewLine & _
        "From (" & strSql & ") A" & vbNewLine & _
        "Group By a.No, a.序号, a.病人id, a.费别, a.收费类别, a.收据费目, a.计算单位, a.开单人," & vbNewLine & _
        "      a.收费细目id, a.保险支付大类id, a.是否医保, a.保险编码" & vbNewLine & _
        "Having Nvl(Sum(a.数量), 0) <> 0" & vbNewLine & _
        "Order By NO, 序号"
    Set rsNo = zlDatabase.OpenSQLRecord(strSql, "获取重新收费数据-医保", strNos)
    
    With rsNo
        p = 0: strNos = ""
        Do While Not rsNo.EOF
            If InStr("," & strNos & ",", "," & Nvl(!NO) & ",") = 0 Then
                strNos = strNos & "," & Nvl(!NO)
                p = p + 1
            End If
            
            rsTmp.AddNew
            rsTmp!单据序号 = p
            rsTmp!费别 = !费别
            rsTmp!NO = Nvl(!NO)
            rsTmp!序号 = Val(Nvl(!序号))
            rsTmp!实际票号 = Nvl(!实际票号)
            rsTmp!结算时间 = !结算时间
            rsTmp!病人ID = Val(Nvl(!病人ID))
            rsTmp!收费类别 = Nvl(!收费类别)
            rsTmp!收据费目 = Nvl(!收据费目)
            rsTmp!开单人 = Nvl(!开单人)
            rsTmp!收费细目ID = Val(Nvl(!收费细目ID))
            rsTmp!计算单位 = Nvl(!计算单位)
            rsTmp!数量 = Val(Nvl(!数量))
            rsTmp!单价 = Val(Nvl(!单价))
            rsTmp!实收金额 = Val(Nvl(!实收金额))
            rsTmp!统筹金额 = Val(Nvl(!统筹金额))
            rsTmp!保险支付大类ID = IIf(Val(Nvl(!保险支付大类ID)) = 0, Null, Val(Nvl(!保险支付大类ID)))
            rsTmp!是否医保 = Val(Nvl(!是否医保))
            rsTmp!保险编码 = Nvl(!保险编码)
            rsTmp!摘要 = Nvl(!摘要)
            rsTmp!是否急诊 = Val(Nvl(!是否急诊))
            rsTmp!开单部门ID = Val(Nvl(!开单部门ID))
            rsTmp!执行部门ID = Val(Nvl(!执行部门ID))
            rsTmp.Update
            
            rsNo.MoveNext
        Loop
    End With
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakePreSwapDataFromDB = rsTmp
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteInsureInfoUpdate(ByVal lng结帐ID As Long, ByVal intInsure As Integer, _
    ByRef objBalanceBills As BalanceBills) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新重收记录的保险信息
    '参数:
    '   str保险金额-"实收合计;进入统筹;全自付;先自付"
    '返回:所有重收记录的保险信息更新成功返回True，否则返回False
    '编制:冉俊明
    '日期:2014-9-16
    '问题:77951
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsReCharge As ADODB.Recordset
    Dim strBXInfo As String, strPreNo As String
    Dim cur实收金额 As Currency, cur统筹金额 As Currency, bln保险项目 As Boolean
    Dim blnTrans As Boolean, cllReChargePro As Collection
    Dim objBalanceBill As BalanceBill
    
    On Error GoTo errHand
    Set objBalanceBills = New BalanceBills
    strSql = _
        "Select a.Id, a.No, a.序号, a.病人id, a.收费细目id, Nvl(a.付数, 1) * a.数次 As 数量, Nvl(a.实收金额, 0) As 实收金额, a.摘要," & vbNewLine & _
        "       Nvl(a.保险项目否, 0) As 保险项目否, a.保险大类id, Nvl(a.统筹金额, 0) As 统筹金额, a.保险编码, a.费用类型" & vbNewLine & _
        "From 门诊费用记录 A" & vbNewLine & _
        "Where Mod(a.记录性质, 10) = 1 And a.结帐id = [1]"
    Set rsReCharge = zlDatabase.OpenSQLRecord(strSql, "获取重收费用记录", lng结帐ID)
    With rsReCharge
        If .RecordCount > 0 Then
            Set cllReChargePro = New Collection
            .Sort = "NO,序号"
            Do While Not .EOF
                If strPreNo <> Nvl(!NO) Then
                    If strPreNo <> "" Then
                        objBalanceBills.AddItem objBalanceBill
                    End If
                    
                    Set objBalanceBill = New BalanceBill
                    objBalanceBill.NO = Nvl(!NO)
                    strPreNo = Nvl(!NO)
                End If
                
                '保险项目否(0/1);保险大类ID;进入统筹金额;保险项目编码;摘要;费用类型
                strBXInfo = gclsInsure.GetItemInsure(Nvl(!病人ID), Nvl(!收费细目ID), Val(Nvl(!实收金额)), True, intInsure, _
                        Nvl(!摘要) & "||" & Val(Nvl(!数量)))
                If strBXInfo <> "" Then
                    '  Zl_门诊收费记录_Update
                    strSql = "Zl_门诊收费记录_Update("
                    '  Id_In         In 门诊费用记录.Id%Type,
                    strSql = strSql & Nvl(!ID) & ","
                    '  保险大类id_In In 门诊费用记录.保险大类id%Type,
                    strSql = strSql & ZVal(Split(strBXInfo, ";")(1)) & ","
                    '  保险项目否_In In 门诊费用记录.保险项目否%Type,
                    strSql = strSql & Val(Split(strBXInfo, ";")(0)) & ","
                    '  保险编码_In   In 门诊费用记录.保险编码%Type,
                    strSql = strSql & "'" & CStr(Split(strBXInfo, ";")(3)) & "',"
                    '  费用类型_In   In 门诊费用记录.费用类型%Type,
                    strSql = strSql & "'" & CStr(Split(strBXInfo, ";")(5)) & "',"
                    '  统筹金额_In   In 门诊费用记录.统筹金额%Type,
                    strSql = strSql & Format(Val(Split(strBXInfo, ";")(2))) & ","
                    '  摘要_In       In 门诊费用记录.摘要%Type
                    strSql = strSql & "'" & CStr(Split(strBXInfo, ";")(4)) & "')"
                    zlAddArray cllReChargePro, strSql
                    
                    cur统筹金额 = CCur(Val(Split(strBXInfo, ";")(2)))
                    bln保险项目 = Val(Split(strBXInfo, ";")(0)) = 1
                Else
                    cur统筹金额 = Val(Nvl(!统筹金额))
                    bln保险项目 = Val(Nvl(!保险项目否)) = 1
                End If
                
                '统计保险金额
                cur实收金额 = Val(Nvl(!实收金额))
                If cur统筹金额 = 0 Or Not bln保险项目 Then
                    '以原始金额为准,不管分币处理
                    objBalanceBill.全自付 = objBalanceBill.全自付 + cur实收金额
                Else
                    objBalanceBill.进入统筹 = objBalanceBill.进入统筹 + cur统筹金额
                    '以原始金额为准,不管分币处理
                    objBalanceBill.先自付 = objBalanceBill.先自付 + cur实收金额 - cur统筹金额
                End If
                objBalanceBill.实收合计 = objBalanceBill.实收合计 + CCur(Val(Nvl(!实收金额)))
                
                rsReCharge.MoveNext
            Loop
            If strPreNo <> "" Then
                objBalanceBills.AddItem objBalanceBill
            End If
            
            '执行过程
            blnTrans = True
            zlExecuteProcedureArrAy cllReChargePro, "执行保险信息更新", True, True
            blnTrans = False
        End If
    End With
    ExecuteInsureInfoUpdate = True
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Function InsureSwapSuccess(ByVal intInsure As Integer, ByVal lng结帐ID As Long) As Boolean
    '医保交易是否成功
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If intInsure = 0 Then InsureSwapSuccess = True: Exit Function
    strSql = _
        "Select 1" & vbNewLine & _
        "From 病人预交记录 A, 保险结算记录 B, 结算方式 C" & vbNewLine & _
        "Where a.结帐id = b.记录id And a.结算方式 = c.名称 And c.性质 In (3, 4) And Nvl(a.校对标志, 0) = 2" & vbNewLine & _
        "      And a.结帐id = [2] And a.卡类别id Is Null And b.险类 = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "判断医保结算是否成功", intInsure, lng结帐ID)
    InsureSwapSuccess = Not rsTemp.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetBalanceDetail(ByVal bytType As Byte, ByVal strValue As String, _
    Optional ByVal bytDataType As Byte, _
    Optional ByVal blnHistory As Boolean) As ADODB.Recordset
    '功能:获取医保结算明细数据
    '入参:
    '   bytType 查找类型:0-根据结帐ID查找;1-根据结算序号查找,2-根据单据号来获取结算方式
    '   strValue 要查找的值(bytType为0时,结帐ID;为1时,结算序号;为2时，为一次收费所涉及的所有单据)
    '   bytDataType 数据类型：1-仅医保结算数据，2-仅一卡通结算数据，0-所有结算数据
    '   bln含异常 根据单据号读取数据时是否读取异常数据
    '返回:返回医保结算明细记录
    '     字段:结帐id,NO,结算方式,金额,卡类别id,关联交易id,交易流水号,交易说明,医保,结算性质
    Dim strSql As String, strWhere As String
    Dim strTable As String
    
    On Error GoTo errHandle
    If bytDataType = 1 Then
        strWhere = " And 卡类别id Is Null"
    ElseIf bytDataType = 2 Then
        strWhere = " And 卡类别id Is Not Null"
    End If
    
    Select Case bytType
    Case 0
        strWhere = strWhere & " And a.结帐ID= [1]"
    Case 1
        strTable = ",(Select Distinct 结帐ID From 病人预交记录 Where 结算序号= [1]) B"
        strWhere = strWhere & " And a.结帐ID = b.结帐ID"
    Case 2
        strTable = _
            ",(Select Distinct 结帐ID  " & _
            "  From 门诊费用记录 " & _
            "  Where Mod(记录性质,10)=1 And NO In (Select Column_value From Table(f_str2List([2]))) And Nvl(费用状态,0)<>1) B"
        strWhere = strWhere & " And a.结帐ID=b.结帐ID"
    End Select
    
    strSql = _
        "Select a.结帐id, a.NO, a.结算方式, a.金额," & vbNewLine & _
        "       a.卡类别id, a.关联交易id, a.交易流水号, a.交易说明," & vbNewLine & _
        "       Decode(c.性质,3,1,4,1,0) As 医保, c.性质 As 结算性质" & vbNewLine & _
        "From 结算方式 C,医保结算明细 A" & strTable & vbNewLine & _
        "Where c.名称 = a.结算方式 " & strWhere
    If blnHistory Then
        strSql = Replace(strSql, "门诊费用记录", "H门诊费用记录")
        strSql = Replace(strSql, "病人预交记录", "H病人预交记录")
        strSql = Replace(strSql, "医保结算明细", "H医保结算明细")
    End If
    
    Set zlGetBalanceDetail = _
        zlDatabase.OpenSQLRecord(strSql, "获取医保结算明细数据", Val(strValue), strValue)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetYBBalanceNo(rsBalance As ADODB.Recordset, Optional ByVal strNos As String, _
    Optional ByVal blnDelCheck As Boolean, Optional ByVal lng病人ID As Long, _
    Optional ByVal intInsure As Integer, Optional ByVal blnDel As Boolean, _
    Optional ByVal bln门诊结算作废 As Boolean, Optional ByVal str个人帐户 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据获取医保原结算方式和结算金额
    '参数：
    '   strNOs - 单据号,多个用逗号隔开：A0001,A0002,...
    '   blnDelCheck - 是否检查允许门诊结算作废
    '返回:返回结算信息,格式:结算方式|结算金额||...
    '编制:刘兴洪
    '日期:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, varNos As Variant, strFilter As String
    Dim i As Integer, p As Integer
    Dim colBalance As Collection, strPreNo As String
    
    On Error GoTo errHandle
    If blnDelCheck And intInsure = 0 Then Exit Function
    If rsBalance Is Nothing Then Exit Function
    
    varNos = Split(strNos, ",")
    For i = 0 To UBound(varNos)
        strFilter = strFilter & " Or No='" & varNos(i) & "'"
    Next
    If strFilter <> "" Then strFilter = Mid(strFilter, 4)
    rsBalance.Filter = strFilter
    If rsBalance.RecordCount = 0 Then Exit Function
    
    '字段:结帐id,NO,结算方式,金额,卡类别id,关联交易id,交易流水号,交易说明,医保,结算性质
    rsBalance.Sort = "No"
    Set colBalance = New Collection
    p = 1: colBalance.Add New BalanceMoneys
    With rsBalance
        strPreNo = Nvl(!NO)
        Do While Not .EOF
            If strPreNo <> Nvl(!NO) Then
                p = p + 1: colBalance.Add New BalanceMoneys
                strPreNo = Nvl(!NO)
            End If
            If blnDelCheck Then
                '如果这种结算方式不支持回退,要退为现金,则不用减去
                If bln门诊结算作废 Then
                    If gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure, !结算方式) Then
                        str结算方式 = Nvl(!结算方式) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(!金额))
                    End If
                Else     '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                    If !结算方式 <> str个人帐户 Then
                        str结算方式 = Nvl(!结算方式) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(!金额))
                    End If
                End If
            Else
                str结算方式 = Nvl(!结算方式) & "|" & Val(Nvl(!金额))
            End If
            
            Call SetBalanceVal(colBalance, p, str结算方式)
            .MoveNext
        Loop
    End With
    GetYBBalanceNo = GetMedicareStr(colBalance)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub CancelBillBalance(ByVal lng结帐ID As Long, Optional ByVal strNO As String, _
    Optional cllPro As Collection)
    '取消单据的结算
    Dim strSql As String
    
    On Error GoTo ErrHandler
    'Zl_门诊收费结算_Cancel_S(
    strSql = "Zl_门诊收费结算_Cancel_S("
    '  结帐id_In   门诊费用记录.结帐id%Type,
    strSql = strSql & "" & lng结帐ID & ","
    '  No_In       门诊费用记录.No%Type := Null
    strSql = strSql & "'" & strNO & "')"
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "取消单据的结算"
    Else
        zlAddArray cllPro, strSql
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function MakeDetailRecord(objBill As BalanceBills) As ADODB.Recordset
'功能：根据单据对象内容创建一个明细记录集信息(以售价单位)
'字段：病人ID，主页ID，收费类别，收费细目ID，数量，单价，实收金额，开单人，开单科室
'参数：intPage=指定的单据,lngRow=指定的行，不指定时包含所有单据的所有行
    Dim i As Integer, p As Integer, strSql As String
    Dim rsTmp As New ADODB.Recordset, rsPrice As ADODB.Recordset
    
    On Error GoTo errHandle
    rsTmp.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "数量", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "单价", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    For p = 1 To objBill.Count
        strSql = "Select a.病人id, a.收费类别, a.收费细目id, Avg(a.数次 * Nvl(a.付数, 1)) As 数量," & vbNewLine & _
                "        Sum(a.标准单价) As 单价, Sum(a.实收金额) 实收金额, a.开单人, b.名称 As 开单科室" & vbNewLine & _
                " From 门诊费用记录 A, 部门表 B" & vbNewLine & _
                " Where a.开单部门id = b.Id And a.No = [1] And a.记录性质 = 1" & vbNewLine & _
                " Group By a.收费细目id, a.病人id, a.收费类别, a.开单人, b.名称"
        Set rsPrice = zlDatabase.OpenSQLRecord(strSql, "读取划价单", objBill(p).NO)
        With rsPrice
            For i = 1 To .RecordCount
                rsTmp.Filter = "收费细目ID=" & !收费细目ID
                If rsTmp.RecordCount = 0 Then
                    rsTmp.AddNew
                    
                    rsTmp!病人ID = !病人ID
                    rsTmp!收费类别 = !收费类别
                    rsTmp!收费细目ID = !收费细目ID
                    
                    rsTmp!数量 = !数量
                    rsTmp!单价 = !单价
                    rsTmp!实收金额 = !实收金额
                    
                    rsTmp!开单人 = !开单人
                    rsTmp!开单科室 = !开单科室
                    
                Else
                    rsTmp!数量 = rsTmp!数量 + !数量
                    rsTmp!单价 = (rsTmp!单价 + !单价) / 2
                    rsTmp!实收金额 = rsTmp!实收金额 + !实收金额
                End If
                
                rsTmp.Update
                .MoveNext
            Next
        End With
    Next
    
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetChargeBalance(ByVal lng结帐ID As Long) As ADODB.Recordset
    '获取结算数据
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim rsTypes As ADODB.Recordset
    On Error GoTo ErrHandler
    
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    strSql = "Select Case" & vbNewLine & _
            "          When Mod(a.记录性质, 10) = 1 Then 1" & vbNewLine & _
            "          When b.名称 Is Not Null And a.卡类别id Is Null Then 2" & vbNewLine & _
            "          When Nvl(a.卡类别id, 0) <> 0 Then 3" & vbNewLine & _
            "          Else 0" & vbNewLine & _
            "        End As 类型, a.Id, Mod(a.记录性质, 10) As 记录性质, a.结算方式, a.冲预交, a.摘要," & vbNewLine & _
            "        a.卡类别id, a.结算卡序号, a.卡号, a.结算号码, a.交易流水号, a.交易说明, a.校对标志," & vbNewLine & _
            "        0 As 是否密文,  '' As 卡类别名称, a.结帐id, a.结算序号" & vbNewLine & _
            " From 病人预交记录 A,  (Select 名称 From 结算方式 Where 性质 In (3, 4)) B" & vbNewLine & _
            " Where a.结算方式 = b.名称(+) And a.结帐ID = [1]" & vbNewLine & _
            "       And (a.记录性质 In (1, 11) Or Nvl(a.结算卡序号, 0) = 0)"
    strSql = strSql & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select 5 As 类型, a.Id, Mod(a.记录性质, 10) As 记录性质, a.结算方式, a.冲预交, a.摘要," & vbNewLine & _
            "        a.卡类别id, a.结算卡序号, a.卡号, a.结算号码, a.交易流水号, a.交易说明, a.校对标志," & vbNewLine & _
            "        Nvl(m.是否密文, 0) As 是否密文, m.名称 As 卡类别名称, a.结帐id, a.结算序号" & vbNewLine & _
            " From 病人预交记录 A, 消费卡类别目录 M" & vbNewLine & _
            " Where a.结算卡序号 = m.编号 And a.记录性质 Not In (1, 11) And a.结帐ID = [1]"
    Set GetChargeBalance = zlDatabase.OpenSQLRecord(strSql, "获取结算数据", lng结帐ID)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetYBBalance(rsBalance As ADODB.Recordset, ByVal lng结帐ID As Long) As String
    '获取医保原结算方式和结算金额
    '返回:返回结算信息,格式:结算方式|结算金额||...
    Dim str结算方式 As String
    
    On Error GoTo errHandle
    rsBalance.Filter = "类型=" & Enum_BalanceType.医保 & " and 结帐ID=" & lng结帐ID
    If rsBalance.RecordCount = 0 Then Exit Function
    
    With rsBalance
        Do While Not .EOF
            str结算方式 = str结算方式 & "||" & Nvl(!结算方式) & "|" & Val(Nvl(!冲预交))
            .MoveNext
        Loop
    End With
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 3)
    GetYBBalance = str结算方式
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetCustomPatiInsure(ByVal lng病人ID As Long) As Integer
    '获取病人险类，在病人识别成功后调用，返回险类后自动调用医保身份识别
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If lng病人ID = 0 Then Exit Function
    '本地如果不支持医保则不调用自定义过程
    If GetSetting("ZLSOFT", "公共全局", "本地支持的医保", "") = "" Then Exit Function
    
    strSql = "Select Zl_Custom_Getpatiinsure([1]) As 险类 From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取病人险类", lng病人ID)
    If rsTemp.EOF Then Exit Function
        
    GetCustomPatiInsure = Val(Nvl(rsTemp!险类))
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
