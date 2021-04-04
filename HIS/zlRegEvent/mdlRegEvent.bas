Attribute VB_Name = "mdlRegEvent"
Option Explicit '要求变量声明
Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
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
    support撤消出院 = 17            '允许撤消病人出院
    support必须录入入出诊断 = 18    '病人入院与出院时，必须录入诊断名
    support记帐完成后上传 = 19      '要求上传在记帐数据提交后再进行
    support出院结算必须出院 = 20    '病人结帐时如果选择出院结帐，就检查必须出院才可以进行
    support挂号使用个人帐户 = 21    '使用医保挂号时是否使用个人帐户进行支付
    
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    support医保接口打印票据 = 46    'HIS中只走票据号但不调打印，医保接口(北京)中打印
    support连续挂号 = 62            '在挂号时，是否允许连继续挂号(在输入缴款金额后才结束)
    support挂号不收取病历费 = 81    '在挂号时，不使用医保收取病历费
    support挂号检查项目 = 86
End Enum
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
Public Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public gobjSquare As SquareCard  '卡结算部件  42301
Public gobjTax As Object '税控打印接口对象
Public gblnTax As Boolean '本机是否使用税控打印
Public gstrTax As String
Public gobjRegist As Object
Public gobjPlugIn As Object, gblnPlugin As Boolean
Public gobjPublicExpense As Object
Public gintPriceGradeStartType As Integer
Public gstrPriceGrade As String

'票据控制
Public gobjBillPrint As Object '第三方票据打印部件
Public gblnBillPrint As Boolean '第三方票据打印部件是否可用


Public Function GetMaxLen() As Byte
'功能：提取挂号项目号别的最大长度
'说明：提取挂号项目编码的最大长度
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    GetMaxLen = 5
    strSQL = "Select Nvl(Max(Length(号码)),5) as 长度 From 挂号安排"
    
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlRegEvent")
    
    If Not rsTmp.BOF Then GetMaxLen = rsTmp!长度
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCashMoney(ByVal strNO As String) As Currency
'功能：医保不支持退个人帐户时,个人帐户退现金,获取现金退款金额
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select -1*A.冲预交 as 现金 From 病人预交记录 A,门诊费用记录 B,结算方式 C " & _
            " Where A.结帐ID=B.结帐ID And A.结算方式=C.名称 And B.NO=[1] " & _
            " And A.记录性质=4 And A.记录状态=2 And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO)
    
    If Not rsTmp.BOF Then GetCashMoney = rsTmp!现金
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get收费执行科室ID(ByVal lng项目id As Long, ByVal int执行科室类型 As Integer) As Long
'功能：获取挂号附加项目(病历费,就诊卡费)的收费项目的执行科室
'参数：
'返回：如果返回零,表示挂号科室(医生所在科室)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Get收费执行科室ID = UserInfo.部门ID
    
    Select Case int执行科室类型
        Case 0 '0-无明确科室
        Case 1 '1-病人所在科室
            Get收费执行科室ID = 0
        Case 2 '2-病人所在病区
            Get收费执行科室ID = 0
        Case 3 '3-操作员科室
        Case 4 '4-指定科室
            strSQL = "Select 执行科室ID From 收费执行科室 Where 收费细目ID=[1] And Nvl(病人来源,1)=1 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", lng项目id)
            
            If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
        Case 5 '院外执行(预留,程序暂未用)
        Case 6 '开单人科室
    End Select
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadRegistPrice(ByVal lng项目id As Long, ByVal bln病历 As Boolean, ByVal bln就诊卡 As Boolean, _
    Optional str费别 As String, Optional rsItems As ADODB.Recordset, Optional rsIncomes As ADODB.Recordset, _
    Optional lng病人ID As Long, Optional int险类 As Integer, Optional str号别 As String, Optional bytMode As Integer, _
    Optional lng挂号科室ID As Long = 0, Optional strPriceGrade As String, Optional strDate As String, _
    Optional ByVal lng卡费细目ID As Long) As Long
'功能：读取指定挂号项目对应的费用信息到记录集中
'参数：lng项目ID=表示是否读取挂号费用(要读的挂号项目ID)
'      bln病历=表示是否读取病历工本费(可能仅收取病历费)
'      bln就诊卡=表示是否读取就诊卡费用(与挂号费或病历费一起收取)
'      str费别=挂号费别
'      rsItems(Out)=包含挂号项目及从属项目,不能以New方式定义
'      rsInComes(Out)=包含各个项目的收入情况,不能以New方式定义
'      strPriceGrade=收费项目的价格等级
'      lng卡费细目ID= 存在自定义卡费时传入
'返回：读取的项目个数,同时rsItems,rsInCome=Nothing
'说明：主项数次为1,从项按设定数次处理,但为固定
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lng原项ID As Long
    Dim rsFeeTmp As ADODB.Recordset
    Dim strFee As String
    Dim str附加项目ID As String
    Dim strWherePriceGrade As String
    Dim strDateCondition As String
    
    Set rsItems = Nothing
    Set rsIncomes = Nothing
    
    If strDate <> "" Then
        strDateCondition = " [5] "
    Else
        strDateCondition = " Sysdate "
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.价格等级 = [4]" & vbNewLine & _
            "          Or (b.价格等级 Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From 收费价目" & vbNewLine & _
            "                             Where b.收费细目id = 收费细目id And 价格等级 = [4]" & vbNewLine & _
            "                                   And " & strDateCondition & " Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null"
    End If
    
    '读取挂号项目及从属项目的费用
    If lng项目id <> 0 Then
        strSQL = _
            "Select 1 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
            " 1 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,-1 as 执行科室类型" & _
            " From 收费项目目录 A,收费价目 B,收入项目 C" & _
            " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=[1]" & _
            " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
        strSQL = strSQL & " Union ALL " & _
            "Select 2 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
            " D.从项数次 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,-1 as 执行科室类型" & _
            " From 收费项目目录 A,收费价目 B,收入项目 C,收费从属项目 D" & _
            " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.从项ID And D.主项ID=[1]" & _
            " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    
    '读取病历工本费对应的费用
    If bln病历 Then
        strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
            "Select 3 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
            " 1 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,A.执行科室 as 执行科室类型" & _
            " From 收费项目目录 A,收费价目 B,收入项目 C,收费特定项目 D" & _
            " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.收费细目ID And D.特定项目='病历费'" & _
            " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    
    '读取就诊卡对应的费用(不支持设置多个收入项目,为了保持和就诊卡管理中一致)
    '变价且最低限价为零时,不收卡费
    If bln就诊卡 Then
        If lng卡费细目ID = 0 Then
            strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
                "Select 4 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
                " 1 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,A.执行科室 as 执行科室类型" & _
                " From 收费项目目录 A,收费价目 B,收入项目 C,收费特定项目 D" & _
                " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.收费细目ID And D.特定项目=[2] And (A.是否变价=1 And Nvl(B.原价,0)<>0 or A.是否变价=0)" & _
                " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD')) And Rownum=1" & vbNewLine & _
                strWherePriceGrade
        Else
            strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
                "Select 4 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
                " 1 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,A.执行科室 as 执行科室类型" & _
                " From 收费项目目录 A,收费价目 B,收入项目 C " & _
                " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=[2] And (A.是否变价=1 And Nvl(B.原价,0)<>0 or A.是否变价=0)" & _
                " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD')) And Rownum=1" & vbNewLine & _
                strWherePriceGrade
        End If
    End If
    
    If bytMode <> 1 And bytMode <> 10 And Not (lng项目id = 0 And bln病历 = True) Then
        strFee = "Select zl_Fun_CustomRegExpenses([1],[2],[3]) As 附加费 From Dual"
        Set rsFeeTmp = zlDatabase.OpenSQLRecord(strFee, "zl_Fun_CustomRegExpenses", lng病人ID, int险类, str号别)
        If Not rsFeeTmp.EOF Then
            str附加项目ID = Nvl(rsFeeTmp!附加费)
        End If
        
        If str附加项目ID <> "" Then
            If strSQL = "" Then
                strSQL = " " & _
                    "Select /*+cardinality(D,10)*/ 5 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
                    " 1 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,-1 as 执行科室类型" & _
                    " From 收费项目目录 A,收费价目 B,收入项目 C,Table(f_str2list([3])) D " & _
                    " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.Column_Value " & _
                    " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
            Else
                strSQL = strSQL & " Union ALL " & _
                    "Select /*+cardinality(D,10)*/ 5 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
                    " 1 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,-1 as 执行科室类型" & _
                    " From 收费项目目录 A,收费价目 B,收入项目 C,Table(f_str2list([3])) D " & _
                    " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.Column_Value " & _
                    " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
            End If
            strSQL = strSQL & " Union ALL " & _
                "Select /*+cardinality(E,10)*/ 5 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
                " D.从项数次 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,-1 as 执行科室类型" & _
                " From 收费项目目录 A,收费价目 B,收入项目 C,收费从属项目 D,Table(f_str2list([3])) E" & _
                " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.从项ID And D.主项ID=E.Column_Value " & _
                " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                strWherePriceGrade
        End If
    End If
    
    If strSQL = "" Then Exit Function
    
    '按主项,从项,病历顺序排列
    strSQL = "Select * From (" & strSQL & ") Order by 性质,项目编码,收入编码"
    
    On Error GoTo errH
    If strDate <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", lng项目id, IIf(lng卡费细目ID = 0, gCurSendCard.str特准项目, lng卡费细目ID), str附加项目ID, strPriceGrade, CDate(strDate))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", lng项目id, IIf(lng卡费细目ID = 0, gCurSendCard.str特准项目, lng卡费细目ID), str附加项目ID, strPriceGrade)
    End If
    
    If Not rsTmp.EOF Then
        '先创建记录集
        Set rsItems = New ADODB.Recordset
        rsItems.Fields.Append "性质", adSmallInt '1-主项,2-从项,3-病历费,4-就诊卡费
        rsItems.Fields.Append "执行科室ID", adBigInt
        rsItems.Fields.Append "类别", adVarChar, 1
        rsItems.Fields.Append "项目ID", adBigInt
        rsItems.Fields.Append "项目名称", adVarChar, 80
        rsItems.Fields.Append "计算单位", adVarChar, 20, adFldIsNullable
        rsItems.Fields.Append "数次", adSingle
        rsItems.Fields.Append "保险项目否", adSmallInt, , adFldIsNullable
        rsItems.Fields.Append "保险大类ID", adBigInt, , adFldIsNullable
        rsItems.Fields.Append "保险编码", adVarChar, 80
        
        rsItems.CursorLocation = adUseClient
        rsItems.LockType = adLockOptimistic
        rsItems.CursorType = adOpenStatic
        rsItems.Open
        
        Set rsIncomes = New ADODB.Recordset
        rsIncomes.Fields.Append "项目ID", adBigInt
        rsIncomes.Fields.Append "收入项目ID", adBigInt
        rsIncomes.Fields.Append "收据费目", adVarChar, 20, adFldIsNullable
        rsIncomes.Fields.Append "单价", adSingle
        rsIncomes.Fields.Append "应收", adCurrency
        rsIncomes.Fields.Append "实收", adCurrency
        rsIncomes.Fields.Append "统筹金额", adCurrency, , adFldIsNullable
        rsIncomes.CursorLocation = adUseClient
        rsIncomes.LockType = adLockOptimistic
        rsIncomes.CursorType = adOpenStatic
        rsIncomes.Open
        
        For i = 1 To rsTmp.RecordCount
            '挂号项目部份
            If lng原项ID <> rsTmp!项目ID Then
                rsItems.AddNew
                rsItems!性质 = rsTmp!性质
                 '0-无明确科室,1-病人所在科室,2-病人所在病区,3-开单人所在科室,4-指定科室
                If bytMode = 10 Then
                    If rsTmp!执行科室类型 = -1 Then
                        rsItems!执行科室ID = lng挂号科室ID
                    Else
                        rsItems!执行科室ID = Get收费执行科室ID(rsTmp!项目ID, rsTmp!执行科室类型)
                    End If
                Else
                    If rsTmp!执行科室类型 = -1 Then
                        rsItems!执行科室ID = 0      '0-表示挂号科室
                    Else
                        rsItems!执行科室ID = Get收费执行科室ID(rsTmp!项目ID, rsTmp!执行科室类型)
                    End If
                End If
                
                rsItems!类别 = rsTmp!类别
                rsItems!项目ID = rsTmp!项目ID
                rsItems!项目名称 = rsTmp!项目名称
                rsItems!计算单位 = rsTmp!计算单位
                rsItems!数次 = Format(Nvl(rsTmp!数次, 0), "0.000")
                rsItems.Update
            End If
            lng原项ID = rsTmp!项目ID
            
            '收入项目部份
            rsIncomes.AddNew
            rsIncomes!项目ID = rsTmp!项目ID
            rsIncomes!收入项目ID = rsTmp!收入项目ID
            rsIncomes!收据费目 = rsTmp!收据费目
            rsIncomes!单价 = Format(Nvl(rsTmp!单价, 0), gstrFeePrecisionFmt)
            rsIncomes!应收 = Format(rsItems!数次 * rsIncomes!单价, "0.00")
            If Nvl(rsTmp!屏蔽费别, 0) = 1 Then
                rsIncomes!实收 = rsIncomes!应收
            Else
                rsIncomes!实收 = Format(GetActualMoney(str费别, rsTmp!收入项目ID, rsIncomes!应收, rsTmp!项目ID), "0.00")
            End If
            rsIncomes.Update
            rsTmp.MoveNext
        Next
        ReadRegistPrice = rsItems.RecordCount
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set rsItems = Nothing
    Set rsIncomes = Nothing
End Function

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim strValue As String
    On Error Resume Next
    
    '上班时间
    strValue = UCase(zlDatabase.GetPara(1, glngSys, , "08:00 AND 12:00"))
    gstr上班时间 = Format(Trim(Split(strValue, "AND")(0)), "HH:mm")
    
    '费用金额小数点位数
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '卡号显示方式
    'gblnShowCard = zlDatabase.GetPara(12, glngSys) = "0"

    '挂号票据号码长度
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbytFactLength = Val(Split(strValue, "|")(IIf(gblnSharedInvoice, 0, 3)))
    'gbyt磁卡 = Val(Split(strValue, "|")(4))
    
    '挂号有效天数
    '刘兴洪:34717
    '两位:前一位普能挂号;后一位急诊挂号
    strValue = zlDatabase.GetPara(21, glngSys, , "01") & "1"
    gSysPara.Sy_Reg.bytNODaysGeneral = Val(Left(strValue, 1))
    gSysPara.Sy_Reg.bytNoDayseMergency = Val(Mid(strValue, 2, 1))
    If gSysPara.Sy_Reg.bytNODaysGeneral = 0 Then gSysPara.Sy_Reg.bytNODaysGeneral = 1
    If gSysPara.Sy_Reg.bytNoDayseMergency = 0 Then gSysPara.Sy_Reg.bytNoDayseMergency = 1
    
    '票号严格控制
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnBill挂号 = (Mid(strValue, IIf(gblnSharedInvoice, 1, 4), 1) = "1")
    'gblnBill磁卡 = (Mid(strValue, 5, 1) = "1")
    
        
    '一卡通消费验证
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    gdbl预存款消费验卡 = Val(Split(strValue, "|")(0))
    gbyt预存款退费验卡 = Val(Split(strValue, "|")(1))
    gbln消费卡退费验卡 = zlDatabase.GetPara(282, glngSys) = "1"
            
    '刷卡要求输入密码
    gstrCardPass = zlDatabase.GetPara(46, glngSys, , "0000000000")
        
    '挂号允许的预约天数
    gint预约天数 = zlDatabase.GetPara(66, glngSys, , 15)
    '刘兴洪 问题:????    日期:2010-12-06 23:38:53
    '费用单价保留位数
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    gbln身份证唯一 = Val(zlDatabase.GetPara("同一身份证只能对应一个建档病人", glngSys)) = 1    '117954
    Call InitAddressLength
    InitSysPar = True
End Function

Public Sub InitLocPar(lngModul As Long)
'功能：初始化本机参数
    Dim strValue As String
    On Error Resume Next
                
    
    'b.数据库存储的公共全局参数
    '----------------------------------------------------------------------------------------
    gstrLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    strValue = zlDatabase.GetPara("输入法")
    gstrIme = IIf(strValue = "", "不自动开启", strValue)
    gbytRegistMode = Val(Split(zlDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(0))
    If Split(zlDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(1) <> "" Then
        gdatRegistTime = CDate(Format(Split(zlDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(1), "yyyy-mm-dd hh:mm:ss"))
    End If
        
    
    'c.数据库存储的模块参数
    '----------------------------------------------------------------------------------------
    If lngModul = 1111 Then
        glngInterval = Val(zlDatabase.GetPara("自动刷新间隔", glngSys, lngModul))
        gbln自动门诊号 = zlDatabase.GetPara("自动门诊号", glngSys, lngModul) = "1"
        gblnPrice = zlDatabase.GetPara("存为划价单", glngSys, lngModul) = "1"
        gblnPrePayPriority = zlDatabase.GetPara("优先使用预交款", glngSys, lngModul) = "1"
        
        '缺省值
        gstr付款方式 = zlDatabase.GetPara("缺省付款方式", glngSys, lngModul)
        gstr费别 = zlDatabase.GetPara("缺省费别", glngSys, lngModul)
        gstr性别 = zlDatabase.GetPara("缺省性别", glngSys, lngModul)
        gstr结算方式 = zlDatabase.GetPara("缺省结算方式", glngSys, lngModul)
        
        '本机允许挂号科室ID
        gstr挂号科室ID = zlDatabase.GetPara("挂号科室", glngSys, lngModul)
        
        
        gblnSeekName = zlDatabase.GetPara("姓名模糊查找", glngSys, lngModul) = "1"
        gintNameDays = Val(zlDatabase.GetPara("姓名查找天数", glngSys, lngModul))
        gbln缴款结束 = zlDatabase.GetPara("缴款挂号结束", glngSys, lngModul) = "1"
        gbln医生 = zlDatabase.GetPara("输入医生", glngSys, lngModul) = "1"
        gblnPrintFree = zlDatabase.GetPara("零费用打印", glngSys, lngModul) = "1"
        gbytInvoice = Val(zlDatabase.GetPara("挂号发票打印方式", glngSys, lngModul, , 1))
        gByt打印病人条码 = Val(zlDatabase.GetPara("病人条码打印方式", glngSys, lngModul, , 1))
        gblnPrintCase = zlDatabase.GetPara("打印病历标签", glngSys, lngModul, "0") = "1"
        gbln精简界面 = Val(zlDatabase.GetPara("计划排班挂号默认界面", glngSys, lngModul, 0)) = 1
        
        
        gbln病人 = zlDatabase.GetPara("输入姓名", glngSys, lngModul) = "1"
        gbln性别 = zlDatabase.GetPara("输入性别", glngSys, lngModul) = "1"
        gbln年龄 = zlDatabase.GetPara("输入年龄", glngSys, lngModul) = "1"
        gbln家庭地址 = zlDatabase.GetPara("输入家庭地址", glngSys, lngModul) = "1"
        gbln付款方式 = zlDatabase.GetPara("输入付款方式", glngSys, lngModul) = "1"
        gbln费别 = zlDatabase.GetPara("输入费别", glngSys, lngModul) = "1"
        gbln结算方式 = zlDatabase.GetPara("输入结算方式", glngSys, lngModul) = "1"
        gbln电话 = zlDatabase.GetPara("输入联系电话", glngSys, lngModul) = "1"
        
        
        gblnAutoAddName = zlDatabase.GetPara("自动产生姓名", glngSys, lngModul) = "1"
        gblnNewCardNoPop = zlDatabase.GetPara("发卡不弹窗口", glngSys, lngModul) = "1"
        gbln卡费仅划价 = zlDatabase.GetPara("收取卡费", glngSys, lngModul) <> "1"
        gbln退费重打 = zlDatabase.GetPara("退费重打", glngSys, lngModul) = "1"
        '问题:35176
        gbyt清除门诊信息 = Val(zlDatabase.GetPara("退号清除门诊信息", glngSys, lngModul))
        
        '收费和挂号共用票据
        gblnSharedInvoice = zlDatabase.GetPara("挂号共用收费票据", glngSys, 1121) = "1"
        '本地共用挂号批次ID
        If gblnSharedInvoice Then
            glng挂号ID = Val(zlDatabase.GetPara("共用收费票据批次", glngSys, 1121, ""))
        Else
            glng挂号ID = Val(zlDatabase.GetPara("共用挂号票据批次", glngSys, lngModul, ""))
        End If
        If glng挂号ID > 0 Then
            If Not ExistBill(glng挂号ID, IIf(gblnSharedInvoice, 1, 4)) Then
                If gblnSharedInvoice Then
                    zlDatabase.SetPara "共用收费票据批次", "0", glngSys, 1121
                Else
                    zlDatabase.SetPara "共用挂号票据批次", "0", glngSys, lngModul
                End If
                glng挂号ID = 0
            End If
        End If
        
        gstr磁卡ID = Val(zlDatabase.GetPara("共用就诊卡批次", glngSys, lngModul, ""))
        
        
        '是否使用LED语音报价器
        gblnLED = Val(GetSetting("ZLSOFT", "公共全局", "使用", 0)) <> 0
        '初始公共发卡信息
        Call InitSendCardPreperty(lngModul)
    ElseIf lngModul = 1114 Then
        Call InitLocVisitPlanPar(1114)
    End If
End Sub

Public Sub InitLocVisitPlanPar(ByVal lngModul As Long)
    '初始化临床出诊的模块参数
    With gVisitPlan_ModulePara
        .byt出诊表打印方式 = Val(zlDatabase.GetPara("出诊表打印方式", glngSys, lngModul, "0"))
        .str号源维护站点 = zlDatabase.GetPara("未区分站点的号源的维护站点", glngSys, lngModul)
        .byt号码比较方式 = Val(zlDatabase.GetPara("号码排序比较方式", glngSys, lngModul))
    End With
End Sub


Public Sub InitSendCardPreperty(ByVal lngModule As Long, Optional lng卡类别ID As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化刷卡属性
    '编制:刘兴洪
    '日期:2011-07-25 11:03:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strSQL As String, blnBoundCard As Boolean
    Dim rsTemp As ADODB.Recordset, str批次 As String, varData As Variant, i As Long
    Dim varTemp  As Variant, ty_Card As Ty_CardProperty
    If lng卡类别ID <> 0 Then
        lngCardTypeID = lng卡类别ID
    Else
        lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, lngModule, 0))
    End If
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '问题号:57326
    strSQL = "" & _
    "   Select Id, 编码, 名称, 短名, 前缀文本, 卡号长度, 缺省标志, 是否固定, 是否严格控制, " & _
    "           nvl(是否自制,0) as 是否自制, nvl(是否存在帐户,0) as 是否存在帐户, " & _
    "           nvl(是否全退,0) as 是否全退,nvl(是否重复使用,0) as 是否重复使用 ,nvl(缺省标志,0) as 缺省标志, " & _
    "           nvl(密码长度,10) as 密码长度,nvl(密码长度限制,0) as 密码长度限制,nvl(密码规则,0) as 密码规则," & _
    "           nvl(是否退现,0) as 是否退现,部件, 备注, 特定项目, 结算方式, 是否启用, 卡号密文," & _
    "           nvl(是否制卡,0) as 是否制卡,nvl(是否发卡,0) as 是否发卡,nvl(是否写卡,0) as 是否写卡, " & _
    "           nvl(发卡性质,0) as 发卡性质, nvl(读卡性质,'1000')  as 读卡性质,nvl(发卡控制,0) as 发卡控制 " & _
    "    From 医疗卡类别 A" & _
    "    Where nvl(是否启用,0)=1 And (ID=[1] " & IIf(lng卡类别ID = 0, "or nvl(缺省标志,0)=1", "") & ")" & _
    "    Order by 编码"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取初始发卡属性", lngCardTypeID)
    If rsTemp.RecordCount >= 2 Then
        rsTemp.Filter = "ID=" & lngCardTypeID
        If rsTemp.EOF Then rsTemp.Filter = 0
    End If
    If rsTemp.RecordCount <> 0 Then
        rsTemp.MoveFirst
        '85565,李南春,2015/7/10:读卡性质
        With ty_Card
            .lng卡类别ID = Val(Nvl(rsTemp!id))
            .str卡名称 = Nvl(rsTemp!名称)
            .lng卡号长度 = Val(Nvl(rsTemp!卡号长度))
            .lng结算方式 = Trim(Nvl(rsTemp!结算方式))
            .bln自制卡 = Val(Nvl(rsTemp!是否自制)) = 1
            .bln严格控制 = Val(Nvl(rsTemp!是否严格控制)) = 1
            .str卡号密文 = Nvl(rsTemp!卡号密文)
            .int密码长度 = Val(Nvl(rsTemp!密码长度))
            .int密码长度限制 = Val(Nvl(rsTemp!密码长度限制))
            .int密码规则 = Val(Nvl(rsTemp!密码规则))
            .bln就诊卡 = .str卡名称 = "就诊卡" And Val(Nvl(rsTemp!是否固定)) = 1
            .str特准项目 = Trim(Nvl(rsTemp!特定项目))
            .bln缺省标志 = Val(Nvl(rsTemp!缺省标志)) = 1
            '问题号:56599
            .bln是否制卡 = Val(Nvl(rsTemp!是否制卡)) = 1
            .bln是否发卡 = Val(Nvl(rsTemp!是否发卡)) = 1
            .bln是否写卡 = Val(Nvl(rsTemp!是否写卡)) = 1
            '问题号:57326
            .lng发卡性质 = Val(Nvl(rsTemp!发卡性质))
            .bln重复使用 = Val(Nvl(rsTemp!是否重复使用)) = 1
            .str读卡性质 = Nvl(rsTemp!读卡性质, "1000")
            .byt发卡控制 = Val(Nvl(rsTemp!发卡控制))
            .blnOneCard = False
            .str短名称 = Nvl(rsTemp!短名)
            If Trim(Nvl(rsTemp!特定项目)) <> "" Then
                Set .rs卡费 = zlGetSpecialItemFee(Trim(Nvl(rsTemp!特定项目)))
                If .bln就诊卡 Then .blnOneCard = GetOneCard.RecordCount > 0
            Else
                Set .rs卡费 = Nothing
            End If
            str批次 = zlDatabase.GetPara("共用医疗卡批次", glngSys, lngModule, "0")
            '领用ID,卡类别ID|...
             .lng共用批次 = 0
            varData = Split(str批次, "|")
            For i = 0 To UBound(varData)
                 varTemp = Split(varData(i), ",")
                 If Val(varTemp(0)) <> 0 Then
                    If Val(varTemp(1)) = .lng卡类别ID Then
                        .lng共用批次 = Val(varTemp(0)): Exit For
                    End If
                 End If
            Next
        End With
    End If
    gCurSendCard = ty_Card
End Sub

Public Function Check发卡性质(lng病人ID As Long, lng卡类别ID As Long, Optional ByVal blnShowMsg As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发卡时检查是否限制病人的发卡张数
    '入参:lng病人ID - 病人ID;lng卡类别ID  - 医疗卡的类别ID
    '     blnShowMsg-是否弹出提示窗
    '编制:王吉
    '问题:57326
    '日期:2013-01-30 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
    strSQL = "Select 名称, 发卡性质 " & _
            "   From 医疗卡类别 A, 病人医疗卡信息 B Where A.ID = B.卡类别ID And B.状态=0 And B.病人ID=[1] And B.卡类别ID=[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "发卡检查", lng病人ID, lng卡类别ID)
    If rsTemp.RecordCount = 0 Then Check发卡性质 = True: Exit Function
    Select Case Val(Nvl(rsTemp!发卡性质, 0))
        Case 0 '不限制
            Check发卡性质 = True
        Case 1 '同一个病人只允许发一张卡
            If blnShowMsg Then
                MsgBox "该病人已经发过" & Nvl(rsTemp!名称) & ",不能在进行发卡操作!", vbInformation + vbOKOnly
            End If
            Check发卡性质 = False
        Case 2 '同一个病人允许发多张卡,但需要提醒
            If blnShowMsg Then
                Check发卡性质 = MsgBox("该病人已经发过" & Nvl(rsTemp!名称) & ",是否要进行发卡操作?", vbQuestion + vbYesNo) = vbYes
            Else
                Check发卡性质 = True
            End If
    End Select
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetNext号别() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Max(号码) as 号码 From 挂号安排 Where Length(号码)=(Select Max(Length(号码)) From 挂号安排)"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlRegEvent")
    
    If Not rsTmp.EOF Then GetNext号别 = zlStr.Increase(IIf(IsNull(rsTmp!号码), "", rsTmp!号码))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get过敏药物(lng病人ID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 过敏药物ID,过敏药物,过敏反应 From 病人过敏药物 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", lng病人ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            Get过敏药物 = Get过敏药物 & ";" & IIf(IsNull(rsTmp!过敏药物ID), "", rsTmp!过敏药物ID) & "|" & IIf(IsNull(rsTmp!过敏药物), "", rsTmp!过敏药物) & "|" & Nvl(rsTmp!过敏反应)
            rsTmp.MoveNext
        Next
        Get过敏药物 = Mid(Get过敏药物, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetField(strSQL As String) As String
'功能：根据SQL语句内容返回第一个字段内容
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlRegEvent")
    If Not rsTmp.EOF Then GetField = IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function RePrintBill(ByVal frmParent As Object, ByVal bytFunc As Byte, _
        ByVal strNO As String, ByVal lng结帐ID As Long, ByVal intInsure As Integer, _
        ByVal blnVirtualPrint As Boolean, _
        Optional strUseType As String, Optional ByVal bln重打 As Boolean, _
        Optional ByVal blnConfirmInvoice As Boolean = True) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:当前收款记录重新打印一张票据
    '入参:
    '   bytFunc:2-退费打印,3-重打,4-补打票据
    '   blnVirtualPrint-医保接口内调用打印，HIS只走票号不实际打印
    '   blnConfirmInvoice:是否需要确认发票号
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-19 17:18:19
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    Dim strInvoice As String
    Dim blnValid As Boolean
    Dim lng领用ID As Long, strBackInvoice As String
    Dim blnReprint As Boolean
    
    '如果严格控制票据使用
    If gblnBill挂号 Then
        If bln重打 Then
            lng领用ID = CheckUsedBill(IIf(gblnSharedInvoice, 1, 4), glng挂号ID, , strUseType)
            Select Case lng领用ID
                Case -1
                    MsgBox "你没有自用和共用的挂号票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            If lng领用ID <= 0 Then Exit Function
        End If
        If bytFunc = 3 Then
            '获取收回票据
            strSQL = _
            "   Select A.号码" & vbNewLine & _
            "   From 票据使用明细 A" & vbNewLine & _
            "   Where A.性质 = 1 And a.原因 <> 6 " & vbNewLine & _
            "           And A.打印id = (Select Max(ID) From 票据打印内容 Where 数据性质 = [2] And NO = [1])" & vbNewLine & _
            "Minus" & vbNewLine & _
            "Select A.号码" & vbNewLine & _
            "From 票据使用明细 A" & vbNewLine & _
            "Where A.性质 = 2 And a.原因 <> 6 " & vbNewLine & _
            "   And A.打印id = (Select Max(ID) From 票据打印内容 Where 数据性质 = [2] And NO = [1])" & vbNewLine & _
            "Order By 号码"
            Set rsInvoice = zlDatabase.OpenSQLRecord(strSQL, "获取收回票据", strNO, 4)
            Do While Not rsInvoice.EOF
                strBackInvoice = strBackInvoice & "," & rsInvoice!号码
                rsInvoice.MoveNext
            Loop
            If strBackInvoice <> "" Then strBackInvoice = Mid(strBackInvoice, 2)
        End If
        blnReprint = bln重打
    End If
    
     '取下一个票据号码
    If Not gblnBill挂号 Then
        If bln重打 = False And bytFunc = 2 Then Exit Function
        '有可能是第一次使用
        Do
            '非严格控制时直接从本地读取
            If gblnSharedInvoice Then
                strInvoice = zlDatabase.GetPara("当前收费票据号", glngSys, 1121)
            Else
                strInvoice = zlDatabase.GetPara("当前挂号票据号", glngSys, 1111)
            End If
            
            If strInvoice = "" Then
                strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定挂号将要使用的开始票据号。" & _
                                vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                "", frmParent.Left + 1500, frmParent.Top + 1500))
            Else
                strInvoice = zlCommFun.IncStr(strInvoice)
                If blnConfirmInvoice Then
                    strInvoice = UCase(InputBox("请确认挂号" & IIf(bytFunc = 4, "补打", "重打") & "使用的开始票据号码：", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                End If
            End If
                
            '用户取消输入,允许打印
            If strInvoice = "" Then
                If MsgBox("你确定不输入挂号票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnValid = True
            Else
                '检查输入有效性
                If zlCommFun.ActualLen(strInvoice) <> gbytFactLength Then
                    MsgBox "输入的挂号票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                    blnConfirmInvoice = True
                Else
                    blnValid = True
                End If
            End If
        Loop While Not blnValid
    Else
        If blnReprint Then
            Do
                '根据票据领用读取
                strInvoice = GetNextBill(lng领用ID)
                If strInvoice = "" Then
                    '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
                    strInvoice = UCase(InputBox("无法根据票据领用情况获取挂号将要使用的开始票据号，" & _
                                    vbCrLf & "请你输入将要使用的开始票据号码：", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                ElseIf blnConfirmInvoice Then
                    strInvoice = UCase(InputBox("请确认挂号" & IIf(bytFunc = 4, "补打", "重打") & "使用的开始票据号码：", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                End If
                
                '用户取消输入,不打印
                If strInvoice = "" Then Exit Function
                
                '检查输入有效性
                If GetInvoiceGroupID(IIf(gblnSharedInvoice, 1, 4), 1, lng领用ID, glng挂号ID, strInvoice, strUseType) = -3 Then
                    MsgBox "你输入的挂号票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
                    blnConfirmInvoice = True
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        Else
            strInvoice = ""
        End If
    End If
    
    '执行数据处理
    Call frmPrint.ReportPrint(bytFunc, strNO, strBackInvoice, lng领用ID, glng挂号ID, strInvoice, _
        zlDatabase.Currentdate, , , , blnVirtualPrint, bytFunc = 2, strUseType)

    RePrintBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub TaxInterface(ByVal byt类型 As Byte, ByVal strPrintNO As String, ByVal strModiNos As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用税控打印接口
    '入参:byt类型-1-正常打印(含修改);2-重打;3-退费
    '        strPrintNO-要打印的单据号，多个时用逗号分隔:'F0000001','F0000002',...
    '        strModiNos-修改多单据中的一张时,指该多张单据的所有NO，用逗号分隔:'F0000001','F0000002',...
    '编制:刘兴洪
    '日期:2013-03-27 14:24:03
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    '未启用税控,直接返回
    If Not gblnTax Then Exit Sub
    If byt类型 = 3 Then
        '退费
        gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strPrintNO, "2")
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        gstrTax = gobjTax.zlTaxOutReput(gcnOracle, strPrintNO, "2")
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If byt类型 = 2 Then
        '重打
        MsgBox "请在准备好之后按确定开始打印。", vbInformation, gstrSysName
        gstrTax = gobjTax.zlTaxOutReput(gcnOracle, strPrintNO, "2")
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If strModiNos <> "" Then
        gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strModiNos, "2")
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
    End If
    gstrTax = gobjTax.zlTaxOutPrint(gcnOracle, strPrintNO, "2")
    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub

Public Function CheckExecuted(strNO As String, blnEnableDel As Boolean) As Boolean
'功能：判断指定的挂号单据是否已经被执行,包括医生接诊下医嘱后作废后,取消接诊,也表示执行过了
'参数:blnEnableDel-是否允许只存在取消过医嘱的病人退号
'返回:True 表示已被执行
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    CheckExecuted = False
    If blnEnableDel Then strSQL = " And 医嘱状态<>4"
    strSQL = _
        " Select count(ID) num From 病人挂号记录 Where NO=[1] And 执行状态>0 and 记录性质=1 and 记录状态 =1 " & _
        " Union All " & _
        " Select count(ID) num From 病人医嘱记录 Where 挂号单=[1] And (病人来源=1 or 病人来源=2)" & strSQL
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO)
    Do While Not rsTmp.EOF
        If rsTmp!Num > 0 Then
            CheckExecuted = True
        End If
        rsTmp.MoveNext
    Loop
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistInsure(strNO As String) As Integer
'功能：判断挂号记录中是否存在指定的医保结算方式
'参数：strNO=挂号单据号
'返回：如果存在则返回单据当时的险类
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select B.险类 From 门诊费用记录 A,保险结算记录 B" & _
       " Where A.记录性质=4 And A.序号=1 And A.记录状态 IN(1,3) And A.NO=[1]" & _
       " And B.性质=1 And A.结帐ID=B.记录ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO)
    
    If Not rsTmp.EOF Then ExistInsure = Val(IIf(IsNull(rsTmp!险类), 0, rsTmp!险类))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function ExistFee(strNO As String) As Boolean
'功能：判断病人的挂号单当天是否还有其它挂号单,如果有,则不检查是否发生过费用,
'      如果没有,则检查是否收过费,不检查划价费用,记帐费用,自动费用,就诊卡费用
'参数：strNO=挂号单据号

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '挂号单当天之后是否有挂号单(不管挂号科室)(如果退号,则病人挂号记录中的数据记录状态不为1)
    strSQL = "Select a.NO, a.病人id, a.执行部门id, a.登记时间,b.执行部门id as 挂号科室id" & vbNewLine & _
            "From 病人挂号记录 a, 病人挂号记录 b" & vbNewLine & _
            "Where b.No = [1] And a.病人id = b.病人id and a.记录性质=1 and a.记录状态=1 and b.记录性质=1 and b.记录状态=1 And a.登记时间 >= Trunc(b.登记时间) and a.记录性质=1 and a.记录状态=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO)
    
    If Not rsTmp.EOF Then
        '如果挂号单的同一科室有多张挂号单,就不检查是否发生过费用了,因为无法区分是哪张挂号单的
        rsTmp.Filter = "执行部门id=" & rsTmp!挂号科室id
        If rsTmp.RecordCount > 1 Then Exit Function
        
        '检查这张挂号单的科室在本次挂号后是否存在费用(未退费)
        rsTmp.Filter = "NO='" & strNO & "'"
        strSQL = "Select NO" & vbNewLine & _
             "From 门诊费用记录" & vbNewLine & _
             "Where 病人id=[1] And 开单部门ID+0=[2] And 登记时间+0>=[3]" & vbNewLine & _
             "      And 记录性质=1 And 记录状态>0 " & vbNewLine & _
             "Group by NO Having Sum(付数*数次)<>0"
             
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", Val(rsTmp!病人ID), Val(rsTmp!执行部门id), CDate(rsTmp!登记时间))
        ExistFee = Not rsTmp.EOF
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPriceHaveFee(strNO As String, ByRef str划价NO As String) As Boolean
'功能:检查挂号产生的划价单是否已经收过费
'返回:未收费的划价单

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select NO,记录状态 From 门诊费用记录 " & _
            " Where 记录性质=1 And 病人ID=(Select 病人ID From 病人挂号记录 Where NO=[1] And 记录性质=1 and 记录状态=1 and  Rownum<2 )" & _
            " And 记录状态 IN(0,1,3) And 序号=1 And 摘要 Like [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO, "%" & strNO & "%")
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!记录状态, 0) = 1 Then
            MsgBox "该挂号单对应的费用已经被单独收费，不能退号。", vbInformation, gstrSysName
            CheckPriceHaveFee = True
        ElseIf Nvl(rsTmp!记录状态, 0) = 0 Then
            str划价NO = rsTmp!NO
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckExistsMCNO(ByVal strMCNO As String) As Boolean
'功能:检查医保号是否已存在
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    strSQL = "Select 1 From 病人信息 Where 医保号 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strMCNO)
    If rsTmp.RecordCount > 0 Then
        MsgBox "请检查,输入的医保号已存在!", vbInformation, gstrSysName
        CheckExistsMCNO = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBill结帐ID(ByVal strNO As String, ByVal byt记录性质 As Byte, _
    Optional ByRef lng病人ID As Long, Optional ByRef bln记帐费用 As Boolean) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据的结帐ID
    '入参:strNo-单据号
    '       byt记录性质:4-挂号,5-就诊卡
    '出参:lng病人ID-返回病人ID
    '       bln记帐费用-返回该单据是否记帐费用
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-19 16:23:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    lng病人ID = 0
    If byt记录性质 <> 5 Then
        strSQL = "Select 结帐ID,病人ID,记帐费用 From 门诊费用记录" & _
           " Where NO=[1] And 记录性质=[2] And 记录状态 IN(1,3) And 序号=1"
    Else
        strSQL = "Select 结帐ID,病人ID,记帐费用 From 住院费用记录" & _
           " Where NO=[1] And 记录性质=[2] And 记录状态 IN(1,3) And 序号=1"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO, byt记录性质)
    If rsTmp.EOF Then Exit Function
    lng病人ID = Val(Nvl(rsTmp!病人ID))
    GetBill结帐ID = Val(Nvl(rsTmp!结帐ID))
    bln记帐费用 = Val(Nvl(rsTmp!记帐费用)) = 1
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckBillExistReplenishData(strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否存在二次结算
    '返回:True-存在二次结算数据 False-不存在二次结算数据
    '入参:strNO-挂号的单据号
    '编制:刘尔旋
    '日期:2014-10-15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSQL = "" & _
    " Select 1" & vbNewLine & _
    " From 费用补充记录 A," & vbNewLine & _
    "     (Select Distinct 结帐id" & vbNewLine & _
    "       From 门诊费用记录" & vbNewLine & _
    "       Where NO = [1] And 记录性质 = 4" & vbNewLine & _
    "       Union" & vbNewLine & _
    "       Select Distinct 结帐id From 住院费用记录 Where NO = [1] And 记录性质 = 5) B" & vbNewLine & _
    " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 1 And Nvl(a.费用状态,0) <> 2 And Rownum < 2"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查二次结算", strNO)

    If rsTmp.EOF Then
        CheckBillExistReplenishData = False
    Else
        CheckBillExistReplenishData = True
    End If
End Function

Public Function GetDoctor(Optional ByVal lngSectID As Long = 0, Optional strCodeAliasName As String = "编码") As ADODB.Recordset
    '得到指定科室下的所有医生并返回
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = _
        "Select c.编号 " & IIf(strCodeAliasName = "", "", " as " & strCodeAliasName) & ",c.姓名,c.简码,c.id From 人员性质说明 a, 部门人员 b ,人员表 c" & vbCrLf & _
        "Where b.人员id=c.id And b.人员id=a.人员id  And  a.人员性质=[1] And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) " & vbCrLf & _
        " And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
        IIf(lngSectID = 0, "", "   And b.部门id = [2]")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", "医生", lngSectID)
    Set GetDoctor = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get医疗付款方式(byt编号 As Byte) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 编码,名称 From 医疗付款方式 Where 编码=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", byt编号)
    If Not rsTmp.EOF Then
        Get医疗付款方式 = rsTmp!编码 & "-" & rsTmp!名称
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDefaultTime(ByVal str号别 As String, vDate As Date) As String
'功能：根据号别在指定日期的预约得出指定日期的缺省时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Decode(" & Weekday(vDate) & ",1,A.周日,2,A.周一,3,A.周二,4,A.周三,5,A.周四,6,A.周五,7,A.周六,NULL)"
    strSQL = "Select B.开始时间 From 挂号安排 A,时间段 B Where " & strSQL & "=B.时间段"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlRegEvent")
    If Not rsTmp.EOF Then
        GetDefaultTime = Format(rsTmp!开始时间, "HH:mm:ss")
    Else
        GetDefaultTime = "00:00:00"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Exist门诊号(str门诊号 As String, Optional lng病人ID As Long) As Boolean
'功能：判断指定门诊号是否已经存在于数据库中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 病人ID From 病人信息 Where 门诊号=[1] And 病人ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", str门诊号, lng病人ID)
    If rsTmp.RecordCount > 0 Then Exist门诊号 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Exist手机号(str手机号 As String, Optional lng病人ID As Long) As Boolean
'功能：判断指定手机号是否已经存在于数据库中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 病人ID From 病人信息 Where 手机号=[1] And 病人ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", str手机号, lng病人ID)
    If rsTmp.RecordCount > 0 Then Exist手机号 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistCardFee(ByVal strNO As String, ByRef lng结帐ID As Long, Optional ByRef strCardNo As String) As String
'功能：判断指定挂号单是否同时收取了就诊卡费
'       返回就诊卡费用单据号,就诊卡费用的结帐ID
'      strCardNo - 卡号
    Dim rsTmp As ADODB.Recordset
    Dim rs医疗卡类别 As Recordset '问题号:56599
    Dim str卡号 As String '问题号:56599
    Dim strSQL As String
    
    On Error GoTo errH
    '问题号:58536
    strSQL = "Select NO,结帐ID,实际票号 as 卡号 From 住院费用记录 Where 记录性质=5 And 记录状态=1 And (病人ID,登记时间) = " & _
            " (Select 病人ID,登记时间 From 门诊费用记录 Where 记录性质=4 And NO=[1] And Rownum=1)"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO)
    If rsTmp.RecordCount > 0 Then
        ExistCardFee = rsTmp!NO
        lng结帐ID = Val(Nvl(rsTmp!结帐ID))
        strCardNo = Nvl(rsTmp!卡号)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check挂号时建档(strNO As String, str挂号日期 As String) As Boolean
    '功能:判断退号时,病人信息是否是挂号时建档的,如果是,则要提示是否清除门诊号
    '由于挂号时发卡新建的病人档案的时间与挂号的时间不一致,以及病人可以先买卡,可能下午或次日再挂号等情况,所以,不能用病人登记时间与挂号时间直接判断
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "" & _
    "   Select 病人id From 病人信息   " & _
    "   Where  ABS(To_Date([2])-登记时间)< (Select Decode(Max(nvl(急诊,0)),0,[3],[4])  From 病人挂号记录 Where NO=[1]  and 记录性质=1 and 记录状态=1) " & _
    "       And 病人ID=(Select 病人id From 病人挂号记录  Where 病人id = (Select 病人id From 病人挂号记录 Where No = [1] and 记录性质=1 and 记录状态=1) and 记录状态=1 and 记录性质=1 " & _
    "   Group By 病人id " & _
    "   Having Count(Id) = 1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO, CDate(str挂号日期), gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency)
    Check挂号时建档 = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SimilarIDs(str身份证号 As String) As String
'功能：检查病人是否存在相似信息
'返回：相似记录的病人ID串,如"234,235,236"
    On Error GoTo errH
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    strSQL = _
        " Select 病人ID,姓名,Nvl(身份证号,'未登记') 身份证号,门诊号,Nvl(家庭地址,'未登记') 地址,To_Char(登记时间,'YYYY-MM-DD') 登记时间 " & _
        " From 病人信息 Where 身份证号=[1]" & _
        " Order by 病人ID Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", str身份证号)
    
    For i = 1 To rsTmp.RecordCount
        SimilarIDs = SimilarIDs & "|ID:" & rsTmp!病人ID & ",姓名:" & rsTmp!姓名 & ",门诊号:" & Nvl(rsTmp!门诊号, "无") & ",身份证号:" & rsTmp!身份证号 & ",地址:" & rsTmp!地址 & ",登记日期:" & rsTmp!登记时间
        rsTmp.MoveNext
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function Select科室(ByVal frmMain As Form, ByVal lngMoudle As Long, ByVal rs科室 As ADODB.Recordset, cbo科室 As ComboBox, ByVal strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择科室
    '入参: frmMain-主窗体
    '      rs科室-传科的科室的本地集,
    '      cbo科室-科室
    '      strKey-选择科室的主键
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-10-12 09:57:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSearch As String
    
    strSearch = "'" & gstrLike & strKey & "%'"
    If IsNumeric(strKey) Then   '输入的是全数字
        rs科室.Filter = "编码 like " & UCase(strSearch)
        rs科室.Sort = "编码"
    ElseIf zlCommFun.IsCharAlpha(strKey) Then  '输入的是全字母
        rs科室.Filter = "简码 like " & UCase(strSearch)
        rs科室.Sort = "简码"
    ElseIf zlCommFun.IsCharChinese(strKey) Then '是否含有汉字,'含有汉字,肯定是找名称
        rs科室.Filter = "名称 like " & strSearch
        rs科室.Sort = "名称"
    Else
        rs科室.Filter = "编码 like " & strSearch & " or 简码 like " & strSearch & " or 名称 like " & strSearch
        rs科室.Sort = "编码"
    End If
    If rs科室.RecordCount = 0 Then
        rs科室.Filter = 0: Exit Function
    End If
    If rs科室.RecordCount = 1 Then
        zlControl.CboLocate cbo科室, Val(Nvl(rs科室!id)), True
        rs科室.Filter = 0: Select科室 = True: Exit Function
    End If
    '弹出选择器
    Dim rsReturn As ADODB.Recordset
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngMoudle, cbo科室, rs科室, True, "", "服务对象", rsReturn) Then
        If Not rsReturn Is Nothing Then
            If rsReturn.RecordCount <> 0 Then
                zlControl.CboLocate cbo科室, Val(Nvl(rsReturn!id)), True
                DoEvents
                If cbo科室.Enabled Then cbo科室.SetFocus
                Select科室 = True: Exit Function
            End If
        End If
    End If
End Function

Public Function zlPersonSelect(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboSel As ComboBox, ByVal rsPerson As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = False, Optional str所有 As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:人员选择选择器
    '入参:cboSel-指定的部门选择部件
    '     rsPerson-指定的人员信息(ID,编号,姓名,简码)
    '     strSearch-要搜索的串
    '     blnNot优先级-是否存在优先级字段
    '     str所有-所有名称(所有人,所有操作员等)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim strIDs As String, str简码 As String, strLike As String
    
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsPerson)
    
    strSearch = UCase(strSearch)
        
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str所有 <> "" Then
        str简码 = zlCommFun.SpellCode(str所有)
        If intInputType = 1 Then
            If Trim(str简码) Like strCompents Then
                rsTemp.AddNew
                rsTemp!id = -1
                rsTemp!编号 = "-"
                rsTemp!姓名 = str所有
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str简码) Like strCompents Or UCase(str所有) Like strCompents Then
                rsTemp.AddNew
                rsTemp!id = -1
                rsTemp!编号 = "-"
                rsTemp!姓名 = str所有
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编号) = strSearch Then lngID = Nvl(!id): iCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编号)) = Val(strSearch) Then
                    If iCount = 0 Then lngID = Val(Nvl(!id))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编号) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!id))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编号) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!姓名)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!id))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编号)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!姓名))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngID = 0
    If lngID <> 0 And rsTemp.RecordCount = 1 Then lngID = Nvl(rsTemp!id)
        
    '刘兴洪:直接定位
    If lngID <> 0 Then GoTo GoOver:
    If lngID < 0 Then lngID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编号"
    Case 1 '输入全拼音
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编号"
    End Select
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboSel, rsTemp, True, "", "缺省," & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngID = Val(Nvl(rsReturn!id))
    If lngID < 0 Then lngID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboSel, lngID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlPersonSelect = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboSel
End Function


Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub

Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '     blnNoBeginTrans:没有事务开始
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub
Public Function zlGetIDCardSex(ByVal strInput As String) As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取身从证号的性别
    '返回：性别
    '编制：刘兴洪
    '日期：2010-07-15 10:31:08
    '说明：15位身份证号码：第7、8位为出生年份(两位数)，第9、10位为出生月份，第11、12位代表出生日期，第15位代表性别，奇数为男，偶数为女。
   '          18位身份证号码：第7、8、9、10位为出生年份(四位数)，第11、第12位为出生月份，第13、14位代表出生日期，第17位代表性别，奇数为男，偶数为女。
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSex As String, i As Integer
    i = zlCommFun.ActualLen(strInput)
    If i <> 15 And i <> 18 Then Exit Function
    i = Val(Mid(strInput, IIf(i = 15, 15, 17), 1))
    If i Mod 2 = 0 Then
        zlGetIDCardSex = "女"
    Else
        zlGetIDCardSex = "男"
    End If
End Function
Public Function zlGetIDCardAge(ByVal strbirthday As Date, ByRef str单位 As String) As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：根据出生日期获取年龄
    '编制：刘兴洪
    '日期：2010-07-15 10:48:15
    '说明：Zl_Age_Calc
    '------------------------------------------------------------------------------------------------------------------------
    Dim lngDiffDay As Long
    If IsDate(strbirthday) = False Then Exit Function
    lngDiffDay = Now - CDate(strbirthday)
    If lngDiffDay < 32 Then '以天为单位
        str单位 = "天": zlGetIDCardAge = lngDiffDay
    ElseIf lngDiffDay < 365 Then
        str单位 = "月": zlGetIDCardAge = Int(lngDiffDay / 30)
    Else
        str单位 = "岁": zlGetIDCardAge = Int(lngDiffDay / 365)
    End If
End Function

Public Sub zlAutoCalcBackLists(ByVal lng病人ID As Long)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：自动计算黑名单
    '编制：刘兴洪
    '日期：2010-07-15 16:32:10
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Err = 0: On Error GoTo Errhand:
    strSQL = "Zl_Regist_Autointoblacklist(" & lng病人ID & ")"
    zlDatabase.ExecuteProcedure strSQL, "计算预约黑名单"
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function GetPatiInfo(lng病人ID As Long) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定的病人信息
    '返回：
    '编制：刘兴洪
    '日期：2010-07-19 10:56:31
    '说明：
    '------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '主页ID=0时(不是NULL)，表示预约入院
    strSQL = _
        " Select A.病人ID,Decode(B.病人ID,NULL,NULL,Nvl(B.主页ID,0)) as 主页ID," & _
        "           A.姓名,A.住院号,B.入院日期,B.出院日期" & _
        " From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID(+) And A.病人ID=[1]" & _
        " Order by Nvl(B.主页ID,0)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng病人ID)
    
    If Not rsTmp.EOF Then Set GetPatiInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlPatiMerge(ByVal lng被合并病人ID As Long, ByRef lng合并病人ID As Long, Optional blnInput合并原因 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：对两个病人信息进行合并
    '入参:blnInput合并原因-是否要求输入合并原因
    '返回：合并成功,返回true, 否则返回False
    '编制：刘兴洪
    '日期：2010-07-19 10:53:12
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, rsPatiS As ADODB.Recordset, rsPatiO As ADODB.Recordset
    Dim strSQL As String, Curdate As Date
    Dim i As Integer, j As Integer
    Dim str合并原因 As String
    
    If lng合并病人ID <= 0 Or lng被合并病人ID <= 0 Then
        Exit Function
    End If
    
    If lng合并病人ID = lng被合并病人ID Then
        MsgBox "相同病人不用进行合并操作！", vbInformation, gstrSysName
        Exit Function
    End If
        
    Set rsPatiS = GetPatiInfo(lng被合并病人ID)
    Set rsPatiO = GetPatiInfo(lng合并病人ID)
    
    'A或B有一个办理了预约入院
    If Not IsNull(rsPatiS!主页ID) And Nvl(rsPatiS!主页ID, 0) = 0 Then
        MsgBox "病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]办理了预约入院登记，请先取消该登记。", vbInformation, gstrSysName
        Exit Function
    End If
    If Not IsNull(rsPatiO!主页ID) And Nvl(rsPatiO!主页ID, 0) = 0 Then
        MsgBox "病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]办理了预约入院登记，请先取消该登记。", vbInformation, gstrSysName
        Exit Function
    End If
    
    'AB都住过院
    If Not IsNull(rsPatiS!主页ID) And Not IsNull(rsPatiO!主页ID) Then
        '1.先住院的在院,不允许(先后住院可以为：出院-出院,出院-在院；不允许：在院-出院,在院-在院)
        '因为除病人合并外,程序不额外处理自动出院或撤消出院
        rsPatiS.MoveLast
        rsPatiO.MoveLast
        If rsPatiS!入院日期 <= rsPatiO!入院日期 Then
            If IsNull(rsPatiS!出院日期) Then
                MsgBox "病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]最后一次住院先入院,但当前未出院,不能执行合并操作！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If IsNull(rsPatiO!出院日期) Then
                MsgBox "病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]最后一次住院先入院,但当前未出院,不能执行合并操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '2.时间交叉提示是否继续
        Curdate = zlDatabase.Currentdate
        rsPatiS.MoveFirst
        For i = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For j = 1 To rsPatiO.RecordCount
                If Not (rsPatiO!入院日期 >= IIf(IsNull(rsPatiS!出院日期), Curdate, rsPatiS!出院日期) Or _
                    IIf(IsNull(rsPatiO!出院日期), Curdate, rsPatiO!出院日期) <= rsPatiS!入院日期) Then
                    MsgBox "发现病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]第 " & rsPatiS!主页ID & " 次住院的期间" & Format(rsPatiS!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiS!出院日期), Curdate, rsPatiS!出院日期), "yyyy-MM-dd") & vbCrLf & _
                        "与病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]的第 " & rsPatiO!主页ID & " 次住院的期间" & Format(rsPatiO!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiO!出院日期), Curdate, rsPatiO!出院日期), "yyyy-MM-dd") & _
                        vbCrLf & "互相交叉，不能进行合并！", _
                        vbInformation, gstrSysName
                        Exit Function
                End If
                rsPatiO.MoveNext
            Next
            rsPatiS.MoveNext
        Next
    End If
    
    '合并原因
    If blnInput合并原因 Then
        str合并原因 = InputBox("合并操作后不能撤消,请慎重!" & vbCrLf & vbCrLf & "请输入合并原因:" & vbCrLf & vbCrLf, gstrSysName, "")
        If zlCommFun.ActualLen(str合并原因) > 250 Then
            MsgBox "合并原因不能多于250个字符,请按Ctrl+C复制下面的内容,重新执行时再输入:" & _
                vbCrLf & vbCrLf & str合并原因, vbInformation, gstrSysName
            Exit Function
        ElseIf Trim(str合并原因) = "" Then
            MsgBox "必须输入合并原因才能进行合并!", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        str合并原因 = "预约挂号接收时,自动合并!"
    End If
    Screen.MousePointer = 11
    DoEvents
    On Error GoTo errH
    strSQL = "zl_病人信息_MERGE(" & lng被合并病人ID & "," & lng合并病人ID & ",'" & str合并原因 & "','" & UserInfo.姓名 & "')"
    
    Call zlDatabase.ExecuteProcedure(strSQL, "自动合并病人")
    On Error GoTo 0
    Screen.MousePointer = 0
        
    '合并后应只剩一个病人
    strSQL = "Select 病人ID From 病人信息 Where 病人ID IN([1],[2])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "自动合并病人", lng被合并病人ID, lng合并病人ID)
    
    lng合并病人ID = Val(rsTmp!病人ID)
    MsgBox "病人合并成功,合并后的病人ID为 " & lng合并病人ID & "。", vbInformation, gstrSysName
    zlPatiMerge = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function



Public Function zl_SelectAndNotAddItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional blnOnlyName As Boolean = False, _
    Optional bln未找到增加 As Boolean = False, Optional strOra过程 As String, Optional strWhere As String, _
    Optional bln站点 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:多功能选择器
    '参数:objCtl-文本框控件
    '     strKey-要搜索的值
    '     strTable-表名
    '     strTittle-选择器名称
    '     bln站点-是否进行站点限制
    '返回:
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str编码 As String, str名称 As String
    Dim vRect As RECT, sngX As Single, sngY As Single, strSQL As String
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    str名称 = strKey
    
    If strTable = "区域" Then
        strSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 And Nvl(级数,0) <3 "
    Else
        strSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    End If
    If strKey <> "" Then
        strSQL = strSQL & _
        "   And ((名称) like [1] or  编码  like [1] or  简码  like  upper([1]))  "
    End If
    strSQL = strSQL & strWhere & IIf(bln站点, zl_获取站点限制, "") & _
    "   order by 编码"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.Hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, strSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If bln未找到增加 Then
            If zlCommFun.IsCharChinese(str名称) = False Then GoTo NOAdd::
            If MsgBox("注意:" & vbCrLf & _
                   "     未找到相关的" & strTable & ",是否增加“" & str名称 & "”？", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If zl_AutoAddBaseItem(strTable, str编码, str名称, strTable & "增加", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = str名称
                    End If
                End With
            Else
                If zlControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                objCtl.Tag = str名称
                zlCommFun.PressKey vbKeyTab
            End If
            zl_SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
            ShowMsgbox "没有找到满足条件的" & strTable & ",请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, Nvl(rsTemp!名称), Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!名称)
            Else
                .Text = IIf(blnOnlyName, Nvl(rsTemp!名称), Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称))
            End If
        End With
    Else
        If zlControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = Nvl(rsTemp!名称)
        objCtl.Tag = Nvl(rsTemp!名称)
        zlCommFun.PressKey vbKeyTab
    End If
    zl_SelectAndNotAddItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zl_获取站点限制(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str别名 As String = "") As String
    '功能:获取站点条件限制:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str别名 = "", "", str别名 & ".") & "站点"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_获取站点限制 = strWhere
End Function
Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.Hwnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub
Public Function zl_AutoAddBaseItem(ByVal strTable As String, str编码 As String, str名称 As String, _
    Optional strTittle As String = "增加项目", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:自动增加项目信息(只针对有编码,名称的信息增加(只增加：编码和名称,简码)
    '--入参数:
    '--出参数:
    '--返  回:增加成功,返回true,否则返回false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, strSQL As String
    Dim int编码 As Integer, strCode As String, strSpecify As String
    zl_AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("没有找到你输入的" & strTable & "，你要把它加入" & strTable & "中吗？", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    
    strSQL = "SELECT Nvl(MAX(LENGTH(编码)), 2) As Length FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, strSQL, strTittle
    
    int编码 = rsTemp!length
    
    strSQL = "SELECT Nvl(MAX(LPAD(编码," & int编码 & ",'0')),'00') As Code FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, strSQL, strTittle
    strCode = rsTemp!Code
    
    int编码 = Len(strCode)
    strCode = strCode + 1
    
    If int编码 >= Len(strCode) Then
    strCode = String(int编码 - Len(strCode), "0") & strCode
    End If
    strSpecify = zlCommFun.SpellCode(str名称)
    
    
    strSQL = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str名称 & "','" & strSpecify & "')"
    zlDatabase.ExecuteProcedure strSQL, strTittle
    str编码 = strCode
    zl_AutoAddBaseItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

 Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitComponents (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Sub
    End If
    
End Sub
Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 关闭结算卡对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         Set gobjSquare.objSquareCard = Nothing
     End If
     Set gobjSquare = Nothing
     If Err <> 0 Then Err.Clear: Err = 0
End Sub

Public Function zl_GetInvoicePrintFormat(ByVal lngModule As Long, Optional str使用类别 As String = "") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票打印格式
    '返回:打印格式(序号)
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
    Dim lngFormat As Long
    Dim lngFormat1 As Long
    
    '因为Getpara就缓存了的,所以不用先用变量进行记录
    strShareTypeFormat = Trim(zlDatabase.GetPara("收费发票格式", glngSys, lngModule, ""))
    '格式:使用类别1,格式1|使用类别2,格式2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lngFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then lngFormat1 = lngFormat
        If Trim(varTemp(0)) = str使用类别 And lngFormat <> 0 Then
            zl_GetInvoicePrintFormat = lngFormat: Exit Function
        End If
    Next
    zl_GetInvoicePrintFormat = lngFormat1
End Function

Public Function zl_GetInvoicePrintMode(ByVal lngModule As Long, _
    Optional str使用类别 As String = "") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票打印格式
    '出参:int打印方式-打印方式()
    '返回:0-不打印;1-自动打印;2-提示打印
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
     Dim intPrintMode As Long, intPrintMode1 As Long
    '因为Getpara就缓存了的,所以不用先用变量进行记录
    strShareTypeFormat = Trim(zlDatabase.GetPara("收费发票打印方式", glngSys, lngModule, ""))
    '格式:使用类别1,打印方式1|使用类别2,打印方式2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",,", ",")
        intPrintMode = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
        If Trim(varTemp(0)) = str使用类别 Then
            zl_GetInvoicePrintMode = intPrintMode: Exit Function
        End If
    Next
    zl_GetInvoicePrintMode = intPrintMode1
End Function

Public Function zl_Get预约方式ByID(lng挂号ID As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据挂号ID获取病人预约方式
    '入参:lng挂号ID-病人挂号ID
    '返回:预约方式
    '编制:王吉
    '日期:2012-07-03
    '问题号:48350
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim str预约方式 As String
    Dim rsTemp As Recordset
    strSQL = "" & _
        "Select 预约方式 From 病人挂号记录 Where 记录状态=1 And ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取预约方式", lng挂号ID)
    If rsTemp Is Nothing Then zl_Get预约方式ByID = "": Exit Function
    If rsTemp.RecordCount = 0 Then zl_Get预约方式ByID = "": Exit Function
    While rsTemp.EOF = False
        str预约方式 = Nvl(rsTemp!预约方式)
        rsTemp.MoveNext
    Wend
    zl_Get预约方式ByID = str预约方式
End Function

Public Function zl_Get预约方式ByNo(strNO As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据挂号单据号获取病人预约方式
    '入参:strNo-挂号单据号
    '返回:预约方式
    '编制:王吉
    '日期:2012-07-03
    '问题号:48350
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim str预约方式 As String
    Dim rsTemp As Recordset
    strSQL = "" & _
        "Select 预约方式 From 病人挂号记录 Where 记录状态=1 And No=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取预约方式", strNO)
    If rsTemp Is Nothing Then zl_Get预约方式ByNo = "": Exit Function
    If rsTemp.RecordCount = 0 Then zl_Get预约方式ByNo = "": Exit Function
    While rsTemp.EOF = False
        str预约方式 = Nvl(rsTemp!预约方式)
        rsTemp.MoveNext
    Wend
    zl_Get预约方式ByNo = str预约方式
End Function

Public Function zl_Get医疗卡类型(lngTypeId As Long) As String()
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据医疗类型ID获取医疗类型
    '入参:lngTypeID-医疗卡类型ID
    '返回:类型对象
    '编制:王吉
    '日期:2012-07-06
    '问题号:51072
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim arr(3) As String
    
    strSQL = "" & _
    "       Select 密码长度,密码输入限制,是否缺省密码 " & _
    "       From 医疗卡类别 " & _
    "       Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取医疗卡类别", lngTypeId)
    If rsTemp Is Nothing Then zl_Get医疗卡类型 = arr: Exit Function
    If rsTemp.RecordCount <= 0 Then zl_Get医疗卡类型 = arr: Exit Function
    rsTemp.MoveFirst
    arr(0) = Nvl(rsTemp!密码长度, "0")
    arr(1) = Nvl(rsTemp!密码输入限制, "0")
    arr(2) = Nvl(rsTemp!是否缺省密码, "0")
    zl_Get医疗卡类型 = arr
End Function
Public Function zlReadRegThreeBalance(ByVal strNO As String, _
    ByRef cllBillBalance As Collection, Optional ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取三方结算交易信息
    '入参:strNo-单据号
    '返回:读取成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2011-08-08 10:10:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng医疗卡类别ID As Long, byt消费卡 As Byte
    Dim objCards As Cards, objTemp As Card
    
    Set cllBillBalance = Nothing
    On Error GoTo errHandle
    '问题:51527: and Mod(B.记录性质,10)<>1"
    gstrSQL = _
        "Select b.结帐id, b.卡类别id, b.结算卡序号, b.卡号, b.交易流水号, b.交易说明, b.合作单位, d.消费卡id" & vbNewLine & _
        "From 门诊费用记录 A, 病人预交记录 B, 病人卡结算记录 D" & vbNewLine & _
        "Where a.结帐id = b.结帐id And b.Id = d.结算id(+) And a.No = [1]" & vbNewLine & _
        "      And a.记录性质 = 4 And a.记录状态 = 1 And Mod(b.记录性质, 10) <> 1" & vbNewLine & _
        "      And (Nvl(b.卡类别id, 0) <> 0 Or Nvl(b.结算卡序号, 0) <> 0)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取结算交易信息", strNO)
    If rsTemp.EOF Then Exit Function
    
    lng医疗卡类别ID = IIf(Val(Nvl(rsTemp!卡类别ID)) > 0, Val(Nvl(rsTemp!卡类别ID)), Val(Nvl(rsTemp!结算卡序号)))
    byt消费卡 = IIf(Val(Nvl(rsTemp!结算卡序号)) <> 0, 1, 0)
    Set objCard = New Card
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO,结帐ID,消费卡ID
    Set cllBillBalance = New Collection
    cllBillBalance.Add Array(lng医疗卡类别ID, Trim(Nvl(rsTemp!卡号)), byt消费卡, _
        Trim(Nvl(rsTemp!交易流水号)), Trim(Nvl(rsTemp!交易说明)), strNO, Val(Nvl(rsTemp!结帐ID)), Val(Nvl(rsTemp!消费卡ID))), strNO
    zlReadRegThreeBalance = True
    If gobjSquare.objSquareCard.zlGetCard(lng医疗卡类别ID, byt消费卡 = 1, objCard) = False Then Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegThreeMoney(lng结帐ID As Long, lngCard结帐ID As Long, _
    ByVal cllBancel As Collection) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取退费的三方交易金额
    '编制:刘兴洪
    '日期:2011-08-08 14:43:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng卡类别ID As Long
    Dim strCardNo As String
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO,结帐ID
    If lng结帐ID = 0 And lngCard结帐ID = 0 Then Exit Function
    If cllBancel Is Nothing Then Exit Function
    lng卡类别ID = Val(cllBancel(1)(0))
    strCardNo = Trim(cllBancel(1)(1))
    strSQL = "Select sum(nvl(冲预交,0)) as 销帐金额 From 病人预交记录 Where 结帐ID in ([1],[2]) and  (卡类别Id=[3] or 结算卡序号=[3]) and mod(记录性质,10)<>1 "
    On Error GoTo Hd
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取三方交易金额", lng结帐ID, lngCard结帐ID, lng卡类别ID, strCardNo)
    If rsTemp.EOF Then Exit Function
    zlGetRegThreeMoney = Val(Nvl(rsTemp!销帐金额))
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Public Function 是否已经签约(strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查需要绑定的卡号是否已经签约
    '入参:绑定卡号
    '编制:王吉
    '日期:2012-08-31 11:32:14
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng身份证类别ID As Long
    Dim rsTemp As Recordset
    On Error GoTo Errhand:
    lng身份证类别ID = Get医疗卡类别ID("二代身份证")
    strSQL = "" & _
    "   Select Count(1) as 是否签约 From 病人医疗卡信息 Where 卡号=[1] And 卡类别ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医疗卡绑定", strCardNo, lng身份证类别ID)
    是否已经签约 = rsTemp!是否签约 > 0
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Public Sub AddSQL绑定卡(ByVal lng病人ID As Long, 卡类别ID As Long, strCard As String, strPassWord As String, ByVal dtCurdate As Date, blnICCard As Boolean, ByRef cllPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绑定卡处理
    '入参:lng病人ID;strCard-绑定卡号;strPassWord-加密密码
    '出参:lngCard结帐ID-卡费的结帐ID
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim str变动原因 As String
    Dim strICCard As String
    
    strICCard = IIf(blnICCard, strCard, "")
    str变动原因 = "病人挂号发卡"
          'Zl_医疗卡变动_Insert
          strSQL = "Zl_医疗卡变动_Insert("
          '      变动类型_In   Number,
          '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
          strSQL = strSQL & "" & 11 & ","
          '      病人id_In     住院费用记录.病人id%Type,
          strSQL = strSQL & "" & lng病人ID & ","
          '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
          strSQL = strSQL & "" & 卡类别ID & ","
          '      原卡号_In     病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'',"
          '      医疗卡号_In   病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      变动原因_In   病人医疗卡变动.变动原因%Type,
          '      --变动原因_In:如果密码调整，变动原因为密码.加密的
          strSQL = strSQL & "'" & str变动原因 & "',"
          '      密码_In       病人信息.卡验证码%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      操作员姓名_In 住院费用记录.操作员姓名%Type,
          strSQL = strSQL & "'" & UserInfo.姓名 & "',"
          '      变动时间_In   住院费用记录.登记时间%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
          strSQL = strSQL & "NULL)"
     zlAddArray cllPro, strSQL
End Sub

Public Function Get医疗卡类别ID(strTypeName As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别ID
    '入参:strTypeName 医疗卡类别名称
    '返回:医疗卡类别ID
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo Errhand
    strSQL = "" & _
    "   Select ID From 医疗卡类别 Where 名称=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医疗卡类别", strTypeName)
    If rsTemp Is Nothing Then Get医疗卡类别ID = 0: Exit Function
    If rsTemp.RecordCount <= 0 Then Get医疗卡类别ID = 0: Exit Function
    Get医疗卡类别ID = rsTemp!id
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatiByID(str类型 As String, strValue As String) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据类型来获取不同条件下的病人信息
    '入参:str类型：查询条件类型 strValue 条件值
    '返回:病人信息集合
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo ErrHandl
    strSQL = "" & _
    "   Select 病人ID,门诊号,住院号,就诊卡号,卡验证码,费别,医疗付款方式,姓名,性别,年龄,出生日期,出生地点,身份证号,其他证件,身份,职业,民族,国籍,籍贯,区域,学历,婚姻状况,家庭地址,家庭电话,家庭地址邮编,监护人," & _
    "   联系人姓名,联系人关系,联系人地址,联系人电话,户口地址,户口地址邮编,Email,QQ,合同单位ID,工作单位,单位电话,单位邮编,单位开户行,单位帐号,担保人,担保性质,就诊时间,就诊状态,就诊诊室,住院次数,当前科室ID,当前床号," & _
    "   入院时间,出院时间,在院,IC卡号,健康号,医保号,险类,查询密码,登记时间,停用时间,锁定,联系人身份证号,结算模式,病人类型,手机号 " & _
    "   From 病人信息 " & _
    "   Where " & str类型 & "=[1]"
    
    Set GetPatiByID = zlDatabase.OpenSQLRecord(strSQL, "门诊挂号", strValue)
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Public Function Get签约病人姓名(str身份证 As String) As String
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:根据身份证获取签约病人的姓名
'入参:str身份证 病人身份证号
'返回:病人姓名
'编制:王吉
'日期:2012-08-31 04:36:33
'问题号:53408
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng身份证类别ID As Long
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl

    lng身份证类别ID = Get医疗卡类别ID("二代身份证")
    strSQL = "" & _
           "   Select 姓名 FROM  病人信息 A,病人医疗卡信息 B Where A.病人ID=B.病人ID And B.卡号=[1] And B.卡类别ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "门诊挂号", str身份证, lng身份证类别ID)
    If rsTemp Is Nothing Then Get签约病人姓名 = "": Exit Function
    If rsTemp.RecordCount Then Get签约病人姓名 = "": Exit Function

    Get签约病人姓名 = rsTemp!姓名
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function


Public Function Bln已发卡(str卡号 As String, lng卡类别 As Long, Optional ByRef lngPatientID As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:获取指定卡号是否已经发卡
'入参:str卡号：卡号 ，lng卡类别：卡类别 , lngPatientID :病人ID
'返回:True :已经发卡;False:未发卡
'编制:王吉
'日期:2012-10-11 04:36:33
'问题号:54390
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl
    strSQL = "" & _
           "   Select 病人ID From 病人医疗卡信息 Where 卡号=[1]  And 卡类别ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "门诊挂号", str卡号, lng卡类别)
    Bln已发卡 = rsTemp.RecordCount > 0

    If rsTemp.RecordCount > 0 Then
        lngPatientID = Val(Nvl(rsTemp!病人ID))
    End If

    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetCardLastChangeType(ByVal str卡号 As String, ByVal lng卡类别 As Long, ByVal lngPaitentID As Long) As Long
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:获取卡最后的变动类型
'入参:str卡号：卡号 ，lng卡类别：卡类别 , lngPatientID :病人ID
'返回:0-未找到相关信息   1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
'编制:李光福
'日期:2013-2-4 17:36:33
'问题号:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "     Select 变动类别" & vbNewLine & _
           "    From (With 医疗卡变动 As (Select 病人id, ID, 变动类别, 变动时间 " & vbNewLine & _
           "                              From 病人医疗卡变动 Bd" & vbNewLine & _
           "                              Where Bd.卡号 = [2] And 卡类别id = [1] And 病人id = [3])" & vbNewLine & _
           "           Select A.变动类别" & vbNewLine & _
           "           From 医疗卡变动 A, (Select Max(变动时间) As 变动时间 From 医疗卡变动 C) B" & vbNewLine & _
           "           Where A.变动时间 = B.变动时间) A"
    On Error GoTo Errhand
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取卡最后变动信息", lng卡类别, str卡号, lngPaitentID)
    If Not rsTmp Is Nothing Then
        If rsTmp.RecordCount > 0 Then
            GetCardLastChangeType = Val(Nvl(rsTmp!变动类别))
        End If
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Public Function zlGetRegAdvanceMoney(lng结帐ID As Long, lngCard结帐ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取退费的预交金额
    '编制:
    '日期:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    If lng结帐ID = 0 And lngCard结帐ID = 0 Then Exit Function
    strSQL = "Select sum(nvl(冲预交,0)) as 销帐金额 From 病人预交记录 Where 结帐ID in ([1],[2]) and mod(记录性质,10)=1 "
    On Error GoTo Hd
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取三方交易金额", lng结帐ID, lngCard结帐ID)
    If rsTemp.EOF Then Exit Function
    zlGetRegAdvanceMoney = Val(Nvl(rsTemp!销帐金额))
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Public Function zlCheckIsAllowBackSN(ByVal strNO As String, _
    ByVal bln记帐 As Boolean, Optional ByRef bln结帐 As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许退号(只检查记帐部分)
    '入参:strNO-退号单据号
    '       bln记帐-是否记帐费用
    '出参:bln结帐-是否需要进行零费用结帐
    '返回:允许退号返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-12-26 09:29:02
    '说明:68991
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng病人ID As Long
    On Error GoTo errHandle
    '只针对记帐费用进行检查
    bln结帐 = False
    If bln记帐 = False Then zlCheckIsAllowBackSN = True: Exit Function
    
    strSQL = " " & _
    "   Select Max(医嘱) As 医嘱, Max(接诊) As 接诊 " & _
    "   From (Select 1 As 医嘱, 0 As 接诊 " & _
    "          From 病人医嘱记录 " & _
    "          Where 挂号单 =[1] " & _
    "          Union All " & _
    "          Select 0 As 医嘱, 1 As 接诊 " & _
    "          From 病人挂号记录 " & _
    "          Where NO = [1] And 记录性质 = 1 And 记录状态 In (1, 3) And 执行状态 In (1, 2))"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查记帐挂号费用的相关状态", strNO)
    
    '1.如果挂号单已经产生了业务数据（即医嘱数据），则也不允许退号
   If Val(Nvl(rsTemp!医嘱)) = 1 Then
        MsgBox "注意:" & vbCrLf & _
                      "       挂号单为" & strNO & "的已经发生了医嘱数据,不允许退号!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
         
    '2.如果挂号单已经接诊,则不允许退号.
    If Val(Nvl(rsTemp!接诊)) = 1 Then
        MsgBox "注意:" & vbCrLf & _
                      "       挂号单为" & strNO & "的已经接诊或完成接诊,不允许退号!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    '3.如果挂号单所对应的记帐单已经结帐，则不允许退号;
    
    strSQL = "" & _
    " Select Nvl(Sum(实收金额), 0) - Nvl(Sum(结帐金额), 0) As 未结金额, Max(结帐id) As 结帐id, Sum(实收金额) As 实收金额,Max(病人ID) as 病人ID " & _
    "   From 门诊费用记录 " & _
    "   Where NO = [1] And Mod(记录性质, 10) = 4 And nvl(记帐费用,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查记帐挂号费用的相关状态", strNO)
    
    If Val(Nvl(rsTemp!未结金额)) = 0 And Val(Nvl(rsTemp!实收金额)) <> 0 Then
        '证明已经结算了
        MsgBox "注意:" & vbCrLf & _
                      "       挂号单为" & strNO & "的已经被结帐,不允许退号!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If Val(Nvl(rsTemp!未结金额)) = 0 And Val(Nvl(rsTemp!实收金额)) = 0 And Val(Nvl(rsTemp!结帐ID)) > 0 Then
        '免费号,实收金额未零,但也可能存在结帐的情况
        MsgBox "注意:" & vbCrLf & _
                      "       挂号单为" & strNO & "的已经被结帐,不允许退号!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
     '4.如果当前挂号单是最后一张有效的挂号单(可能存在多个号,检查标准：挂号有效天数)，病人还存在记帐数据时，也不允许退号
    lng病人ID = Val(Nvl(rsTemp!病人ID))

    strSQL = " " & _
    "   Select Count(*)  as 个数,Max(当前单据) as 当前单据 " & _
    "   From ( Select  distinct NO,decode(NO,[2],1,0) as 当前单据 " & _
    "               From 门诊费用记录 " & _
    "               Where 病人id = [1] And 记录性质 = 4 And 记录状态 = 1 And " & _
    "                         (       (Nvl(加班标志, 0) = 1 And 登记时间+0 >= Sysdate-" & gSysPara.Sy_Reg.bytNoDayseMergency & ")  " & _
    "                            Or (Nvl(加班标志, 0) = 0 And 登记时间+0 >= Sysdate -" & gSysPara.Sy_Reg.bytNODaysGeneral & "))) "
   Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查记帐挂号费用的相关状态", lng病人ID, strNO)
    If Val(Nvl(rsTemp!个数)) = 1 Then
        If Val(Nvl(rsTemp!当前单据)) = 1 Then
            '最后一张单据,需要检查是否有记帐数据
            strSQL = "" & _
            "   Select  sum(金额) as 金额 " & _
            "   From 病人未结费用 " & _
            "   Where 病人id = [1] And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 " & _
            "   UNION ALL " & _
            "   Select -1*Sum(实收金额 ) From 门诊费用记录 " & _
            "   Where No=[2] and 记录性质=4 and 记录状态=1 and nvl(记帐费用,0)=1 "
            strSQL = "Select sum(金额) as 金额 From (" & strSQL & ")"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查记帐挂号费用的相关状态", lng病人ID, strNO)
            
            If Val(Nvl(rsTemp!金额)) <> 0 Then
                MsgBox "注意:" & vbCrLf & _
                "       挂号单为" & strNO & "的已经被结帐,不允许退号!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            '退挂号有效期外的单据,暂不处理
            bln结帐 = True
        End If
    End If
    zlCheckIsAllowBackSN = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str病人类型 As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人类型,设置不同病人类型的显示颜色
    '入参:objPatiControl-病人控件(文本框,标签)
    '    str病人类型-病人类型
    '    lngDefaultColor-缺省病人的显示颜色
    '返回:True-设置颜色成功，False-失败
    '编制:李南春
    '日期:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str病人类型 <> "" Then
        lngColor = zlDatabase.GetPatiColor(str病人类型)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function

Public Function CheckStructAddr(ByVal objCtl As PatiAddress, ByVal lngLen As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查结构化地址控件中的信息录入是否正确
    '入参:objCtl-结构化地址控件，lngLen-限制长度
    '返回:True-输入信息合法
    '编制:李南春
    '日期:2015-12-7
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If zlCommFun.ActualLen(objCtl.Value) > lngLen Then
        MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "最多只能输入" & lngLen \ 2 & "个汉字,请检查。", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    If objCtl.CheckNullValue(, True, False) <> "" Then
        MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "的" & objCtl.CheckNullValue & "尚未输入,请检查。", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    CheckStructAddr = True
End Function

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'功能：外挂创建与检查
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Function ExcPlugInFun(ByVal bytFunc As Byte, ByVal lngRegID As Long, Optional strDoctor As String, Optional strRoom As String, _
                                Optional strNewArrange As String, Optional lngNewArrangeID As Long) As Boolean
    '功能:挂号分诊检测接口
    'bytFunc - 0-分诊;1-换号;2-完成就诊(13-恢复就诊);3-标记为不就诊;4-签道(14-取消签道);5-回诊(15-取消回诊);6-病人待诊
    If gblnPlugin = False Then ExcPlugInFun = True: Exit Function
    
    On Error Resume Next
    ExcPlugInFun = gobjPlugIn.PatiRegTriageCheck(glngSys, glngModul, bytFunc, lngRegID, strDoctor, strRoom, strNewArrange, lngNewArrangeID)
    If Err.Number <> 0 Then
        Call zlPlugInErrH(Err, "PatiRegTriageCheck")
        Err.Clear
        ExcPlugInFun = True
    End If
    On Error GoTo 0
End Function

Public Function Get项目金额(ByVal lng项目id As Long, ByVal strPriceGrade As String) As Double
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWherePriceGrade As String
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.价格等级 = [2]" & vbNewLine & _
            "          Or (b.价格等级 Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From 收费价目" & vbNewLine & _
            "                             Where b.收费细目id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
            "                                   And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null"
    End If
    strSQL = "Select 项目id, Sum(Nvl(现价, 0)) As 金额" & vbNewLine & _
            "From (Select /*+cardinality(D,10)*/" & vbNewLine & _
            "        b.现价, d.Column_Value As 项目id" & vbNewLine & _
            "       From 收费项目目录 A, 收费价目 B, 收入项目 C, Table(f_Str2list([1])) D" & vbNewLine & _
            "       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.Column_Value And Sysdate Between b.执行日期 And" & vbNewLine & _
            "             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select /*+cardinality(E,10)*/" & vbNewLine & _
            "        b.现价 * d.从项数次, e.Column_Value As 项目id" & vbNewLine & _
            "       From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D, Table(f_Str2list([1])) E" & vbNewLine & _
            "       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And d.主项id = e.Column_Value And Sysdate Between b.执行日期 And" & vbNewLine & _
            "             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')))" & vbNewLine & _
            "Group By 项目id"
    
    strSQL = "" & vbNewLine & _
            "Select b.现价" & vbNewLine & _
            "       From 收费项目目录 A, 收费价目 B, 收入项目 C" & vbNewLine & _
            "       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = [1] And Sysdate Between b.执行日期 And" & vbNewLine & _
            "             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
    strSQL = strSQL & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select b.现价 * d.从项数次" & vbNewLine & _
            "       From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D" & vbNewLine & _
            "       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And d.主项id = [1] And Sysdate Between b.执行日期 And" & vbNewLine & _
            "             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
    strSQL = "Select Sum(Nvl(现价, 0)) As 金额 From (" & strSQL & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取挂号项目金额", lng项目id, strPriceGrade)
    If rsTemp.EOF Then
        Get项目金额 = 0
    Else
        Get项目金额 = Val(Nvl(rsTemp!金额))
    End If
End Function

Public Function Get项目信息(ByVal str项目ids As String, ByVal strPriceGrade As String) As ADODB.Recordset
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWherePriceGrade As String
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.价格等级 = [2]" & vbNewLine & _
            "          Or (b.价格等级 Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From 收费价目" & vbNewLine & _
            "                             Where b.收费细目id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
            "                                   And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null"
    End If
    strSQL = "Select /*+cardinality(D,10)*/" & vbNewLine & _
            "       b.现价, d.Column_Value As 项目id" & vbNewLine & _
            " From 收费项目目录 A, 收费价目 B, 收入项目 C, Table(f_Str2list([1])) D" & vbNewLine & _
            " Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.Column_Value And Sysdate Between b.执行日期 And" & vbNewLine & _
            "       Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
    strSQL = strSQL & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select /*+cardinality(E,10)*/" & vbNewLine & _
            "        b.现价 * d.从项数次, e.Column_Value As 项目id" & vbNewLine & _
            " From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D, Table(f_Str2list([1])) E" & vbNewLine & _
            " Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And d.主项id = e.Column_Value And Sysdate Between b.执行日期 And" & vbNewLine & _
            "       Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
    strSQL = "Select 项目id, Sum(Nvl(现价, 0)) As 金额" & vbNewLine & _
            " From (" & strSQL & ")" & vbNewLine & _
            " Group By 项目id"
    Set Get项目信息 = zlDatabase.OpenSQLRecord(strSQL, "获取挂号项目金额", str项目ids, strPriceGrade)
End Function

Public Function RoundEx(ByVal dblNumber As Double, ByVal intBit As Integer) As Double
'功能：四舍五入方式格式化数字
'参数：intBit=最大小数位数
'问题号：94552
'说明：VB自带的Round是银行家舍入法,与实际不一致。如Round(57.575,2)=57.58,Round(57.565,2)=57.56
    If intBit > 0 Then
        RoundEx = Val(Format(dblNumber, "0." & String(intBit, "0")))
    Else
        RoundEx = dblNumber
    End If
End Function

Public Sub CreatePublicExpenseObject(ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共费用部件
    '入参:
    '编制:
    '日期:
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjPublicExpense Is Nothing Then
        Set gobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If gobjPublicExpense Is Nothing Then Exit Sub
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If gobjPublicExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)初始化失败，请与系统管理员联系！", vbExclamation, gstrSysName
         Exit Sub
    End If
    
    gintPriceGradeStartType = gobjPublicExpense.zlGetPriceGradeStartType()
    If gintPriceGradeStartType = 0 Then Exit Sub
    '读取站点价格等级
    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", , , gstrPriceGrade)
End Sub

Public Function ZlGetBillFormat(ByVal intFormat As Integer) As String
    '功能：获取票据格式名称
    '入参：
    '   intFormat - 票据格式序号
    '返回：票据格式的名称
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strRptName As String
    
    On Error GoTo ErrHandler
    strRptName = "ZL" & glngSys \ 100 & "_BILL_1111"
    
    If intFormat = 0 Then '以缺省票据格式显示
        intFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
    End If
    
    strSQL = _
        "Select b.说明" & vbNewLine & _
        "From zlReports A, zlRPTFMTs B" & vbNewLine & _
        "Where a.Id = b.报表id And a.编号 = [1] And b.序号 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取报表格式说明", strRptName, intFormat)
    If rsTmp.EOF Then Exit Function
    
    ZlGetBillFormat = Nvl(rsTmp!说明)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function CreateRegisterObject() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建费用公共的挂号对象
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-04 10:06:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zlPublicExpense.clsRegist")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense.clsRegist)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    
    If gobjRegist Is Nothing Then Exit Function
    'zlInitCommon(ByVal lngSys As Long,  ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    If gobjRegist.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense.clsRegist)初始化失败，请与系统管理员联系！", vbExclamation, gstrSysName
         Exit Function
    End If
    CreateRegisterObject = True
End Function

Public Function CheckBillRepeat(ByVal lng领用ID As Long, ByVal byt票种 As Byte, ByVal strFactNO As String) As Boolean
'功能：在使用新票号之前，检查是否重复
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    strSQL = _
        "Select 号码" & vbNewLine & _
        "From 票据使用明细" & vbNewLine & _
        "Where 领用ID = [1] And 票种=[2] And 号码=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng领用ID, byt票种, strFactNO)
    CheckBillRepeat = Not rsTmp.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PatiValiedCheckByPlugIn(ByVal lngModule As Long, _
    ByVal lng病人ID As Long, ByVal strPatiInforXML As String) As Boolean
    '调用外挂接口 PatiValiedCheck 检查病人信息
    '问题号:102230,106686,138602
    '说明：
    '   1.没有外挂部件时，认为检查通过
    '   2.外挂部件中无PatiValiedCheck接口，也认为检查通过
    '   3.建档病人在识别病人成功后调用，未建档病人在保存数据前调用
    
    If CreatePlugInOK(lngModule) = False Then PatiValiedCheckByPlugIn = True: Exit Function
    If gobjPlugIn Is Nothing Then PatiValiedCheckByPlugIn = True: Exit Function
    
    On Error Resume Next
    'PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, _
        ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, _
        ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
    '功能：检查当前病人是否是指定的特殊病人
    '返回：true时允许继续操作，False时不允许操作
    '参数：
    '      lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '      lngType 操作类型：1－门诊挂号，2－住院入院，3－门诊收费，4－住院结帐。
    '      lngPatiID-病人ID: 新建档的，为0,否则传入建档病人ID
    '      lngPageID-主页ID: 新建档的，为0,否则传入建档主页ID(住院传入主页ID) 特殊说明：仅 lngType=4 时才传入 lngPageID，其它均传0
    '      strPatiInforXML-病人信息:针对未建档病人传入，"姓名，性别，年龄，出生日期，医保号，身份证号，医生姓名"，出生日期 格式:2016-11-11 12:12:12
    '                      固定格式：<XM></XM><XB></XB><NL></NL><CSRQ></CSRQ><YBH></YBH><SFZH></SFZH><YSXM></YSXM>
    '                   建档病人传入，"医生姓名"(106686),格式：<YSXM></YSXM>
    '      strReserve=保留参数,用于扩展使用
    If gobjPlugIn.PatiValiedCheck(glngSys, lngModule, 1, lng病人ID, 0, strPatiInforXML) = False Then
        '注意，接口不存在时也会进入
        If Err <> 0 Then
            If Err.Number = 438 Then '接口不存在，认为检查通过
                PatiValiedCheckByPlugIn = True: Exit Function
            End If
            Call zlPlugInErrH(Err, "PatiValiedCheck")
        End If
        Exit Function
    End If
    PatiValiedCheckByPlugIn = True
End Function

Public Sub InitAddressLength()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select 家庭地址, 户口地址, 出生地点, 联系人地址 From 病人信息 Where Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取地址长度")
    If Not rsTmp.EOF Then
        glngMax家庭地址 = rsTmp.Fields("家庭地址").DefinedSize
        glngMax户口地址 = rsTmp.Fields("户口地址").DefinedSize
        glngMax出生地点 = rsTmp.Fields("出生地点").DefinedSize
        glngMax联系人地址 = rsTmp.Fields("联系人地址").DefinedSize
    End If
    If glngMax家庭地址 = 0 Then glngMax家庭地址 = 100: If glngMax户口地址 = 0 Then glngMax户口地址 = 100
    If glngMax出生地点 = 0 Then glngMax出生地点 = 100: If glngMax联系人地址 = 0 Then glngMax联系人地址 = 100
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function Get分诊科室(ByVal lngSys As Long, ByVal lngModul As Long, ByVal strPrivs As String) As String
    Dim rsTmp As ADODB.Recordset
    Dim str分诊科室 As String, strTmp As String
    On Error GoTo errH
    str分诊科室 = zlDatabase.GetPara("分诊科室", lngSys, lngModul)
    If InStr(strPrivs, "所有科室") = 0 Then
        Set rsTmp = GetDepartments("'临床'", "1,3", InStr(strPrivs, "所有科室") = 0)
        If str分诊科室 = "" Then
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & Nvl(rsTmp!id)
                rsTmp.MoveNext
            Loop
        Else
            Do While Not rsTmp.EOF
                If InStr("," & str分诊科室 & ",", "," & Nvl(rsTmp!id) & ",") > 0 Then
                    strTmp = strTmp & "," & Nvl(rsTmp!id)
                End If
                rsTmp.MoveNext
            Loop
        End If
        If strTmp <> "" Then
            str分诊科室 = Mid(strTmp, 2)
        Else
            str分诊科室 = "0"
        End If
    End If
    Get分诊科室 = str分诊科室
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Get分诊科室 = "0"
End Function
