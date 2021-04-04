Attribute VB_Name = "mdlPubFormCode"
Option Explicit
'*********************************************************************************************************************************************
'模块说明:多个挂号窗体通用规则
'1.控件耦合：确保控件变量与窗体中的控件名称一致，方便查找
'   1)GetPayFromList:将支付列表中的数据更新到支付对象集合中
'   2)AddPayToList:将支付数据更新到支付列表中
'   3)zlGetBalanceSQLByVsf:挂号时从列表中获取未结算的sql
'   4)SelectMemo:选择挂号摘要
'   5)SetDelMemo:选择退号摘要
'   6)Load现住址:加载现住址
'   7)SetTxtTop:改变文本框定点但不引起字体自适应变化
'       7.1)SetTxtLeft
'       7.2)SetTxtWidth
'2.通用规则
'   1)GetPayInfo:更加选择的支付方式获取支付信息
'   2)zlIsAllowPatiChargeFeeMode:检查是否允许改变病人收费模式
'   3)zlCheckBackCard:检查退号时的退卡操作是否合法
'   4)zlGetBackInvoice-获取退号发票
'   5)zlGetBalanceInfor:获取退号结算信息
'   6)GetAll医生-获取医生列表
'   7)zlCheckBackCard-退号时的退卡检查
'   8)GetPatiIDByComminuty-根据社区号获取病人ID
'   9)GetColItem-获取集合中指定的节点值
'   10)GetRoom-获取挂号的诊室
'       10.1)GetRoomVisit
'  12)zlGet失约号-获取安排在某一天.预约失约数
'3.外挂函数
'   1)Plug_PatiValiedCheck:检查病人是否有效，无效则禁止挂号
'*********************************************************************************************************************************************
Private mrsDoctor As ADODB.Recordset
'控件耦合
'*********************************************************************************************************************************************
Public Function GetPayFromList(objRegInfor As clsRegEventInfor, ByVal vsfPay As VSFlexGrid) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将支付列表中数据保存到对象集合
    '入参:objRegInfor-挂号信息、vsfPay支付列表
    '返回:
    '编制:李南春
    '日期:2019/1/29 9:42:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim dbl支付金额 As Double
    Dim lng卡类别id As Long, bln消费卡 As Boolean
    Dim objPay As clsPayInfo, objPayReg As clsSubPayInfo, objCard As Card
    Dim str结算方式 As String, strRows As String   '已经汇总统计了的行数
    
    On Error GoTo Errhand
    objRegInfor.objPayInfos.ReMoveAll
    With vsfPay
        For i = 1 To vsfPay.Rows - 1
            If InStr(strRows & ",", "," & i & ",") = 0 And .RowHidden(i) = False And (.RowData(i) <> 11 Or Val(.TextMatrix(i, .ColIndex("金额"))) > 0) Then
                strRows = strRows & "," & i
                dbl支付金额 = Val(.TextMatrix(i, .ColIndex("金额"))) - Val(.TextMatrix(i, .ColIndex("已支付")))
                lng卡类别id = Val(.TextMatrix(i, .ColIndex("卡类别ID")))
                bln消费卡 = Val(.TextMatrix(i, .ColIndex("消费卡"))) = 1
                str结算方式 = .TextMatrix(i, .ColIndex("结算方式"))
                '如果三方卡仅支付卡费,则不存值
                If Not (RoundEx(dbl支付金额, 6) <= RoundEx(objRegInfor.objPayInfos.Card_结算金额, 6) And _
                   lng卡类别id = objRegInfor.objPayInfos.Card_卡类别ID And _
                   bln消费卡 = objRegInfor.objPayInfos.Card_消费卡 And _
                   str结算方式 = objRegInfor.objPayInfos.Card_结算方式 And _
                   (objRegInfor.Card_变动类型 = CP_发卡 Or objRegInfor.Card_变动类型 = CP_退卡)) Then
                   
                    Set objPay = New clsPayInfo 'clsPayInfo
                    objPay.名称 = .TextMatrix(i, .ColIndex("支付方式"))
                    
                    If lng卡类别id = objRegInfor.objPayInfos.Card_卡类别ID And _
                        bln消费卡 = objRegInfor.objPayInfos.Card_消费卡 And _
                        str结算方式 = objRegInfor.objPayInfos.Card_结算方式 And _
                        (objRegInfor.Card_变动类型 = CP_发卡 Or objRegInfor.Card_变动类型 = CP_退卡) Then
                        dbl支付金额 = dbl支付金额 - objRegInfor.objPayInfos.Card_结算金额
                        objRegInfor.objPayInfos.Card_校对标志 = Val(.TextMatrix(i, .ColIndex("校对标志")))
                        objRegInfor.objPayInfos.Card_结算成功 = Val(.Cell(flexcpData, i, .ColIndex("校对标志"))) = 1
                    End If
                    objPay.支付金额 = dbl支付金额
                    objPay.结算方式 = str结算方式
                    objPay.结算性质 = .RowData(i)
                    objPay.接口序号 = lng卡类别id
                    objPay.消费卡 = bln消费卡
                    objPay.消费卡ID = Val(.TextMatrix(i, .ColIndex("消费卡ID")))
                    objPay.卡号 = .TextMatrix(i, .ColIndex("卡号"))
                    objPay.校对标志 = Val(.TextMatrix(i, .ColIndex("校对标志")))
                    objPay.结算成功 = Val(.Cell(flexcpData, i, .ColIndex("校对标志"))) = 1 '固定=1
                    objPay.交易流水号 = .TextMatrix(i, .ColIndex("交易流水号"))
                    objPay.交易说明 = .TextMatrix(i, .ColIndex("交易说明"))
                    objPay.关联交易ID = Val(.TextMatrix(i, .ColIndex("关联交易ID")))
                    objPay.PayRow = i
                    
                    
                    Set objPayReg = New clsSubPayInfo 'clsSubPayInfo
                    objPayReg.PayRow = i
                    objPayReg.结算方式 = .TextMatrix(i, .ColIndex("结算方式"))
                    objPayReg.结算金额 = dbl支付金额
                    objPayReg.结算号码 = .TextMatrix(i, .ColIndex("结算号码"))
                    objPayReg.交易流水号 = .TextMatrix(i, .ColIndex("交易流水号"))
                    objPayReg.交易说明 = .TextMatrix(i, .ColIndex("交易说明"))
                    objPay.AddItem objPayReg, "K" & objPayReg.PayRow
                    
                    If objPay.接口序号 > 0 Then
                        For j = i + 1 To vsfPay.Rows - 1
                            If objPay.接口序号 = Val(.TextMatrix(j, .ColIndex("卡类别ID"))) And objPay.消费卡 = (Val(.TextMatrix(j, .ColIndex("消费卡"))) = 1) Then
                                str结算方式 = .TextMatrix(i, .ColIndex("结算方式"))
                                dbl支付金额 = Val(.TextMatrix(j, .ColIndex("金额"))) - Val(.TextMatrix(j, .ColIndex("已支付")))
                                If lng卡类别id = objRegInfor.objPayInfos.Card_卡类别ID And _
                                    bln消费卡 = objRegInfor.objPayInfos.Card_消费卡 And _
                                    str结算方式 = objRegInfor.objPayInfos.Card_结算方式 And _
                                    (objRegInfor.Card_变动类型 = CP_发卡 Or objRegInfor.Card_变动类型 = CP_退卡) Then
                                    dbl支付金额 = dbl支付金额 - objRegInfor.objPayInfos.Card_结算金额
                                End If
                                
                                strRows = strRows & "," & j
                                Set objPayReg = New clsSubPayInfo 'clsSubPayInfo
                                objPayReg.PayRow = j
                                objPayReg.结算方式 = .TextMatrix(j, .ColIndex("结算方式"))
                                objPayReg.结算金额 = dbl支付金额
                                objPayReg.结算号码 = .TextMatrix(j, .ColIndex("结算号码"))
                                objPayReg.交易流水号 = .TextMatrix(j, .ColIndex("交易流水号"))
                                objPayReg.交易说明 = .TextMatrix(j, .ColIndex("交易说明"))
                                objPay.AddItem objPayReg, "K" & objPayReg.PayRow
                            End If
                        Next
                    ElseIf .RowData(i) = 3 Or .RowData(i) = 4 Then '医保支持多种结算方式，也按相同的方式保存
                        For j = i + 1 To vsfPay.Rows - 1
                            If (.RowData(j) = 3 Or .RowData(j) = 4) And Val(.TextMatrix(j, .ColIndex("卡类别ID"))) = 0 Then
                                dbl支付金额 = Val(.TextMatrix(j, .ColIndex("金额"))) - Val(.TextMatrix(j, .ColIndex("已支付")))
                                strRows = strRows & "," & j
                                Set objPayReg = New clsSubPayInfo 'clsSubPayInfo
                                objPayReg.PayRow = j
                                objPayReg.结算方式 = .TextMatrix(j, .ColIndex("结算方式"))
                                objPayReg.结算金额 = dbl支付金额
                                objPayReg.结算号码 = .TextMatrix(j, .ColIndex("结算号码"))
                                objPayReg.交易流水号 = .TextMatrix(j, .ColIndex("交易流水号"))
                                objPayReg.交易说明 = .TextMatrix(j, .ColIndex("交易说明"))
                                objPay.AddItem objPayReg, "K" & objPayReg.PayRow
                            End If
                        Next
                    End If
                    
                    If objPay.接口序号 <> 0 Then ' 现在不用区分消费
                        If objPay.消费卡 Then
                            objPay.支付类型 = Pay_SquarePay
                        Else
                            objPay.支付类型 = Pay_ThreePay
                        End If
                        objRegInfor.objPayInfos.AddItem objPay, IIf(objPay.消费卡, "X", "K") & objPay.接口序号
                    Else
                        objPay.支付类型 = Decode(vsfPay.RowData(i), 11, Pay_AccountPay, 2, Pay_CashPay, 3, Pay_InsurePay, 4, Pay_InsurePay, Pay_CashPay)
                        objRegInfor.objPayInfos.AddItem objPay, "PAY" & objPay.支付类型 & "_" & objPay.PayRow
                    End If
                    If objPay.支付类型 = Pay_AccountPay Then
                        objRegInfor.objPayInfos.预交金 = objPay.支付金额
                    End If
                Else
                    objRegInfor.objPayInfos.Card_校对标志 = Val(.TextMatrix(i, .ColIndex("校对标志")))
                    objRegInfor.objPayInfos.Card_结算成功 = Val(.Cell(flexcpData, i, .ColIndex("校对标志"))) = 1
                    objRegInfor.objPayInfos.Card_PayRow = i
                End If
            End If
        Next
    End With
    GetPayFromList = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function UpdatePayToList(objPayInfos As clsPayInfos, ByVal vsfPay As VSFlexGrid) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 跟新列表中的结算方式,主要用于退款成功后更新校对标志=2
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/3/23 09:13
    '---------------------------------------------------------------------------------------
    Dim objPay As clsPayInfo, objSubPay As clsSubPayInfo
    On Error GoTo Errhand
    For Each objPay In objPayInfos
        If objPay.Count > 0 Then
            For Each objSubPay In objPay
                If objSubPay.PayRow > 0 Then
                    With vsfPay
                        .TextMatrix(objSubPay.PayRow, .ColIndex("校对标志")) = objPay.校对标志
                        If objPay.校对标志 = 2 Then
                            .Cell(flexcpForeColor, objSubPay.PayRow, 0, objSubPay.PayRow, .Cols - 1) = 0
                        End If
                    End With
                End If
            Next
        Else
            If objPay.PayRow > 0 Then
                With vsfPay
                    .TextMatrix(objPay.PayRow, .ColIndex("校对标志")) = objPay.校对标志
                    If objPay.校对标志 = 2 Then
                        .Cell(flexcpForeColor, objPay.PayRow, 0, objPay.PayRow, .Cols - 1) = 0
                    End If
                End With
            End If
        End If
    Next
    If objPayInfos.Card_PayRow > 0 Then
        With vsfPay
            .TextMatrix(objPayInfos.Card_PayRow, .ColIndex("校对标志")) = objPayInfos.Card_校对标志
            If objPayInfos.Card_校对标志 = 2 Then
                .Cell(flexcpForeColor, objPayInfos.Card_PayRow, 0, objPayInfos.Card_PayRow, .Cols - 1) = 0
            End If
        End With
    End If
    UpdatePayToList = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AddPayToList(ByVal objPay As clsPayInfo, ByVal vsfPay As VSFlexGrid, _
                Optional ByVal bln异常重收 As Boolean, Optional ByVal byt支付类型 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将结算信息更新到支付列表中
    '入参:objPayInfo-结算信息
    '     byt支付类型 - 0-未实际支付，在最后一步完成；1-已支付完成；2-支付正在进行中，未明确结果
    '返回:
    '编制:李南春
    '日期:2019/1/29 9:42:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim objSubPay As clsSubPayInfo
    Dim str结算方式 As String
    
    If objPay Is Nothing Then AddPayToList = True: Exit Function
    If objPay.支付金额 = 0 Then AddPayToList = True: Exit Function
    With vsfPay
        If objPay.Count > 0 Then
            .RemoveItem objPay.PayRow
            For Each objSubPay In objPay
                If str结算方式 <> objSubPay.结算方式 Then str结算方式 = objSubPay.结算方式: .Rows = .Rows + 1
                .RowData(.Rows - 1) = objPay.结算性质
                .TextMatrix(.Rows - 1, .ColIndex("支付方式")) = objSubPay.结算方式
                .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format(Val(.TextMatrix(.Rows - 1, .ColIndex("金额"))) + objSubPay.结算金额, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = objSubPay.结算方式
                .TextMatrix(.Rows - 1, .ColIndex("结算号码")) = objSubPay.结算号码
                .TextMatrix(.Rows - 1, .ColIndex("卡类别ID")) = objPay.接口序号
                .TextMatrix(.Rows - 1, .ColIndex("消费卡")) = IIf(objPay.消费卡, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("消费卡ID")) = objPay.消费卡ID
                .TextMatrix(.Rows - 1, .ColIndex("卡号")) = objPay.卡号
                .TextMatrix(.Rows - 1, .ColIndex("关联交易ID")) = objPay.关联交易ID
                .TextMatrix(.Rows - 1, .ColIndex("修改")) = IIf(byt支付类型 = 0, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("金额修改")) = IIf(byt支付类型 = 0, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("校对标志")) = IIf(byt支付类型 = 1, 2, 1)
                .Cell(flexcpData, .Rows - 1, .ColIndex("校对标志")) = IIf(byt支付类型 = 1, 1, 0) '固定
                .TextMatrix(.Rows - 1, .ColIndex("独立结算")) = IIf(objPay.独立结算, 1, 0)
                If byt支付类型 = 2 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                Else
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = 0
                End If
            Next
        Else
            '如果指定列无效，则插入到最后一列中
            If objPay.PayRow = 0 Or objPay.PayRow > .Rows - 1 Then
                objPay.PayRow = 0
                For lngRow = 1 To .Rows - 1
                    If Trim(.TextMatrix(lngRow, .ColIndex("支付方式"))) = "" Then
                        objPay.PayRow = lngRow: Exit For
                    End If
                Next
                If objPay.PayRow = 0 Then
                    objPay.PayRow = .Rows
                    .Rows = .Rows + 1
                End If
            ElseIf bln异常重收 Then
                .RemoveItem objPay.PayRow
                objPay.PayRow = .Rows
                .Rows = .Rows + 1
            End If
            .RowData(objPay.PayRow) = objPay.结算性质
            .TextMatrix(objPay.PayRow, .ColIndex("支付方式")) = objPay.名称
            .TextMatrix(objPay.PayRow, .ColIndex("金额")) = Format(Val(.TextMatrix(objPay.PayRow, .ColIndex("金额"))) + objPay.支付金额, "0.00")
            .TextMatrix(objPay.PayRow, .ColIndex("结算方式")) = objPay.结算方式
            .TextMatrix(objPay.PayRow, .ColIndex("结算号码")) = objPay.结算号码
            .TextMatrix(objPay.PayRow, .ColIndex("卡类别ID")) = objPay.接口序号
            .TextMatrix(objPay.PayRow, .ColIndex("消费卡")) = IIf(objPay.消费卡, 1, 0)
            .TextMatrix(objPay.PayRow, .ColIndex("消费卡ID")) = objPay.消费卡ID
            .TextMatrix(objPay.PayRow, .ColIndex("卡号")) = objPay.卡号
            .TextMatrix(objPay.PayRow, .ColIndex("关联交易ID")) = objPay.关联交易ID
            .TextMatrix(objPay.PayRow, .ColIndex("修改")) = IIf(byt支付类型 = 1, 0, 1)
            .TextMatrix(objPay.PayRow, .ColIndex("金额修改")) = IIf(byt支付类型 = 0, 1, 0)
            .TextMatrix(objPay.PayRow, .ColIndex("校对标志")) = IIf(byt支付类型 = 1, 2, 1)
            .Cell(flexcpData, objPay.PayRow, .ColIndex("校对标志")) = IIf(byt支付类型 = 1, 1, 0) '固定
            .TextMatrix(objPay.PayRow, .ColIndex("独立结算")) = IIf(objPay.独立结算, 1, 0)
            If byt支付类型 = 2 Then
                .Cell(flexcpForeColor, objPay.PayRow, 0, objPay.PayRow, .Cols - 1) = vbRed
            Else
                .Cell(flexcpForeColor, objPay.PayRow, 0, objPay.PayRow, .Cols - 1) = 0
            End If
        End If
    End With
End Function

Public Function zlGetBalanceSQLByVsf(ByVal objRegInfor As clsRegEventInfor, ByVal vsfPay As VSFlexGrid, _
                                    cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 :
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/2/26 16:29
    '---------------------------------------------------------------------------------------
    Dim dbl金额 As Double
    Dim str结算方式 As String, str结算信息 As String, strSQL As String
    Dim PayType As gPagePay
    Dim i As Long
    On Error GoTo errH
    If cllPro Is Nothing Then Set cllPro = New Collection
    With vsfPay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("校对标志"))) = 1 And (Val(.TextMatrix(i, .ColIndex("消费卡"))) = 1 Or Val(.TextMatrix(i, .ColIndex("卡类别ID"))) = 0) Then
                dbl金额 = Val(.TextMatrix(i, .ColIndex("金额"))) - Val(.TextMatrix(i, .ColIndex("已支付")))
                str结算方式 = .TextMatrix(i, .ColIndex("结算方式"))
                If str结算方式 = objRegInfor.objPayInfos.Card_结算方式 And objRegInfor.objPayInfos.Card_结算金额 <> 0 Then
                    dbl金额 = dbl金额 - objRegInfor.objPayInfos.Card_结算金额
                    strSQL = zlGetCardFeeModifySQL(False, objRegInfor.objPayInfos.Card_单据号, objRegInfor.objPayInfos.Card_结帐ID, str结算方式, _
                                objRegInfor.objPayInfos.Card_结算金额, , , Val(.TextMatrix(i, .ColIndex("卡类别ID"))), _
                                IIf(Val(.TextMatrix(i, .ColIndex("消费卡"))) = 1, True, False), .TextMatrix(i, .ColIndex("卡号")), , , , .TextMatrix(i, .ColIndex("结算号码")))
                    Call zlAddArray(cllPro, strSQL)
                End If
                If RoundEx(dbl金额, 6) > 0 Then
                    If Val(.TextMatrix(i, .ColIndex("消费卡"))) = 1 Then
                        str结算信息 = str结算方式 & "," & dbl金额
                        PayType = Pay_SquarePay
                    Else
                        str结算信息 = str结算方式 & "," & dbl金额 & "," & .TextMatrix(i, .ColIndex("结算号码")) & ", "
                        PayType = Pay_CashPay
                    End If
                    strSQL = zlGetRegFeeModifySQL(False, objRegInfor.objPayInfos.Reg_单据号, objRegInfor.objPayInfos.Reg_结帐ID, str结算信息, PayType, , , , , _
                                Val(.TextMatrix(i, .ColIndex("卡类别ID"))), .TextMatrix(i, .ColIndex("卡号")))
                    Call zlAddArray(cllPro, strSQL)
                End If
            End If
        Next
    End With
'    '卡费为0时单独处理,缺省为现金，结算方式可能不在列表中
'    If objRegInfor.Card_变动类型 = CP_发卡 And objRegInfor.objPayInfos.Card_结算金额 = 0 Then
'        strSql = zlGetCardFeeModifySQL(False, objRegInfor.objPayInfos.Card_单据号, objRegInfor.objPayInfos.Card_结帐ID, objRegInfor.objPayInfos.Card_结算方式, 0)
'        Call zlAddArray(cllPro, strSql)
'    End If
    zlGetBalanceSQLByVsf = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SelectMemo(frmMain As Form, cbo备注 As ComboBox, ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择常用摘要
    '入参:strInput-输入串;为空时,表示全部
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-04 16:06:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSQL As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If zlCommFun.IsCharChinese(cbo备注.Text) Then
             strWhere = " And  名称 like [1] "
        ElseIf zlCommFun.IsNumOrChar(cbo备注.Text) Then
             strWhere = " And (简码 like upper([1]) or 编码 like upper([1]))"
        End If
    End If
    
    strSQL = "" & _
     "   Select RowNum AS ID,编码,名称,简码  " & _
     "   From 常用挂号摘要 " & _
     "   Where 1=1 " & strWhere & _
     "   Order by 缺省标志"
     vRect = zlControl.GetControlRect(cbo备注.Hwnd)
     On Error GoTo Hd
     Set rsInfo = zlDatabase.ShowSQLSelect(frmMain, strSQL, 0, "常用挂号摘要", False, _
                    "", "", False, False, True, vRect.Left, vRect.Top, cbo备注.Height, blnCancel, True, False, strKey)
     If blnCancel Then Exit Function
     If rsInfo Is Nothing Then
        If strInput = "" Then
            MsgBox "没有设置常用挂号摘要,请在字典管理中设置", vbInformation, gstrSysName
        End If
        zlCommFun.PressKey vbKeyTab: Exit Function
     End If
     Call zlControl.CboSetText(cbo备注, Nvl(rsInfo!名称))
     cbo备注.Tag = Nvl(rsInfo!名称)
     If cbo备注.Visible And cbo备注.Enabled Then cbo备注.SetFocus
     zlCommFun.PressKey vbKeyTab
     SelectMemo = True
     Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Public Function SetDelMemo(ByVal cbo备注 As ComboBox, ByVal strInput As String) As Boolean
    Dim rsMemo As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    cbo备注.Clear
    If strInput = "" Then
        strSQL = "Select 名称,缺省标志 From 常用退号原因 Order By 缺省标志 Desc,编码"
        Set rsMemo = zlDatabase.OpenSQLRecord(strSQL, "SetDelMemo")
        If rsMemo.RecordCount <> 0 Then
            Do While Not rsMemo.EOF
                cbo备注.AddItem rsMemo!名称
                If Val(Nvl(rsMemo!缺省标志)) = 1 Then
                    '不触发Click事件
                    Call cbo.SetIndex(cbo备注.Hwnd, cbo备注.NewIndex): cbo备注.Tag = cbo备注.Text
                End If
                rsMemo.MoveNext
            Loop
        End If
    Else
        strSQL = "Select 名称,缺省标志,简码,编码 From 常用退号原因 Order By 缺省标志 Desc,编码"
        Set rsMemo = zlDatabase.OpenSQLRecord(strSQL, "SetDelMemo")
        If rsMemo.RecordCount <> 0 Then
            Do While Not rsMemo.EOF
                cbo备注.AddItem rsMemo!名称
                If Nvl(rsMemo!简码) Like UCase(strInput) & "*" Or Nvl(rsMemo!编码) Like UCase(strInput) & "*" Or Nvl(rsMemo!名称) Like strInput & "*" Then
                    '不触发Click事件
                    Call cbo.SetIndex(cbo备注.Hwnd, cbo备注.NewIndex): cbo备注.Tag = cbo备注.Text
                End If
                rsMemo.MoveNext
            Loop
            If cbo备注.Text = "" Then
                MsgBox "没有找到对应的退号原因,请重新输入", vbInformation, gstrSysName
                SetDelMemo = False
                Exit Function
            End If
        End If
    End If
    SetDelMemo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Public Function Load现住址(ByVal frmMain As Form) As ADODB.Recordset
    Dim strSQL As String, strFile As String
    Dim fld As Field, rsCheck As ADODB.Recordset
    Dim fso As Scripting.FileSystemObject
    Dim rsCopy As ADODB.Recordset
    Dim rsNew As ADODB.Recordset, rs现住址 As New ADODB.Recordset
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\ZLAddressForRegEvent.Adtg"
    
    On Error Resume Next
    If fso.FileExists(strFile) Then
        rs现住址.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '仅Update时才锁定
        If rs现住址.RecordCount > 0 Then
            rs现住址!次数 = rs现住址!次数 + 1
        End If
        If Err <> 0 Then
            rs现住址.Close
        Else
            rs现住址!次数 = rs现住址!次数 - 1
        End If
    End If
    Err.Clear: On Error GoTo errH
    
    If rs现住址.State = 0 Then
        strSQL = "Select '系统' As 类别, 名称, 简码, 1 As 次数 From 地区"
        Set rsCopy = zlDatabase.OpenSQLRecord(strSQL, "Load现住址")     '必须是adUseClient才能建索引
        Set rs现住址 = zlDatabase.CopyNewRec(rsCopy)
        If Not rs现住址.EOF Then
            '创建索引:名称,简码
            Set fld = rs现住址.Fields(1)
            fld.Properties("Optimize") = True
            Set fld = rs现住址.Fields(2)
            fld.Properties("Optimize") = True
            
            If fso.FileExists(strFile) Then
                Kill strFile
            End If
            rs现住址.Save strFile, adPersistADTG
        End If
        rs现住址.Close
        rs现住址.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '仅Update时才锁定
    Else
        strSQL = "Select '系统' As 类别, 名称, 简码, 1 As 次数 From 地区 Where 1 = 0"
        Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "Load现住址")
        If rsCheck.Fields(1).DefinedSize > rs现住址.Fields(1).DefinedSize Or rsCheck.Fields(2).DefinedSize > rs现住址.Fields(2).DefinedSize Then
            If fso.FileExists(strFile) Then
                Kill strFile
            End If
            strSQL = "Select '系统' As 类别, 名称, 简码, 1 As 次数 From 地区"
            Set rsCopy = zlDatabase.OpenSQLRecord(strSQL, "Load现住址")
            Set rsNew = zlDatabase.CopyNewRec(rsCopy)
            rsNew.Save strFile, adPersistXML
            rs现住址.Close
            rs现住址.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '仅Update时才锁定
        End If
    End If
    
    frmMain.lbl现住址.ToolTipText = "请定期备份本机[病人现住址]数据文件:" & strFile
    Set Load现住址 = rs现住址
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set Load现住址 = rs现住址
End Function

Public Function SetTxtTop(ByVal txtThis As TextBox, ByVal lngTop As Long)
    '---------------------------------------------------------------------------------------
    ' 功能 : 调整控件位置，但不引起控件根据字体的自适应变化
    ' 入参 : txtThis-需要调整的控件
    '        lngLeft-左边距离
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/5/17 13:52
    '---------------------------------------------------------------------------------------
    Dim objFont As StdFont
    On Error Resume Next
    Set objFont = New StdFont
    objFont.Size = txtThis.FontSize
    
    txtThis.FontSize = 1
    txtThis.Top = lngTop
    
    txtThis.FontSize = objFont.Size
End Function

Public Function SetTxtLeft(ByVal txtThis As TextBox, ByVal lngLeft As Long)
    '---------------------------------------------------------------------------------------
    ' 功能 : 调整控件位置，但不引起控件根据字体的自适应变化
    ' 入参 : txtThis-需要调整的控件
    '        lngLeft-左边距离
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/5/17 13:52
    '---------------------------------------------------------------------------------------
    Dim objFont As StdFont
    On Error Resume Next
    Set objFont = New StdFont
    objFont.Size = txtThis.FontSize
    
    txtThis.FontSize = 1
    txtThis.Left = lngLeft
    
    txtThis.FontSize = objFont.Size
End Function

Public Function SetTxtWidth(ByVal txtThis As TextBox, ByVal lngWidth As Long)
    '---------------------------------------------------------------------------------------
    ' 功能 : 调整控件位置，但不引起控件根据字体的自适应变化
    ' 入参 : txtThis-需要调整的控件
    '        lngWidth-宽度
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/5/17 13:52
    '---------------------------------------------------------------------------------------
    Dim objFont As StdFont
    On Error Resume Next
    Set objFont = New StdFont
    objFont.Size = txtThis.FontSize
    
    txtThis.FontSize = 1
    txtThis.Width = lngWidth
    
    txtThis.FontSize = objFont.Size
End Function

'通用规则
'*********************************************************************************************************************************************
Public Function GetPayInfo(ByVal colCardPayMode As Collection, ByVal str结算方式 As String, _
                            objPayInfo As clsPayInfo) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据结算名称获取结算信息
    ' 入参 : colCardPayMode:支付信息集合，窗体加载支付方式时初始化
    '      : str结算方式 :需要获取的结算方式
    ' 出参 : objPayInfo：包括结算性质、结算性质、结算方式、接口序号、是否消费卡、支付类型、是否独立结算
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2018/11/20 09:30
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    If objPayInfo Is Nothing Then Set objPayInfo = New clsPayInfo
    If str结算方式 = "" Then GetPayInfo = True: Exit Function
    
    
    If str结算方式 = "预交金" Or str结算方式 = "退预交" Then
        objPayInfo.名称 = str结算方式
        objPayInfo.支付类型 = Pay_AccountPay
        objPayInfo.结算性质 = 11
        GetPayInfo = True: Exit Function
    End If
    
    '优先充集合中查找,但存在一个问题，医疗卡名称和消费卡名称相同
    'colCardPayMode:名称，性质，结算方式，卡类别ID，是否消费卡，是否独立结算
    For i = 1 To colCardPayMode.Count
        If colCardPayMode(i)(0) = str结算方式 Then
            objPayInfo.名称 = str结算方式
            objPayInfo.结算性质 = Val(colCardPayMode(i)(1))
            objPayInfo.结算方式 = colCardPayMode(i)(2)
            objPayInfo.接口序号 = Val(colCardPayMode(i)(3))
            objPayInfo.消费卡 = Val(colCardPayMode(i)(4)) = 1
            objPayInfo.独立结算 = Val(colCardPayMode(i)(5)) = 1
            If objPayInfo.接口序号 > 0 Then
                If objPayInfo.消费卡 Then
                    objPayInfo.支付类型 = Pay_SquarePay
                Else
                    objPayInfo.支付类型 = Pay_ThreePay
                End If
            Else
                objPayInfo.支付类型 = Pay_CashPay
            End If
            GetPayInfo = True
            Exit Function
        End If
    Next
    ' 什么时候选中了结算方式但又是缓存中没有的
    MsgBox "无效的结算方式", vbInformation, gstrSysName
    Exit Function
    
    strSQL = "Select 性质 From 结算方式 Where 名称=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetPayInfo", str结算方式)
    If Not rsTemp.EOF Then
        objPayInfo.名称 = str结算方式
        objPayInfo.结算性质 = Val(rsTemp!性质)
        objPayInfo.结算方式 = objPayInfo.名称
    Else 'nexttodo
'        strSql = "Select 1 As 类型, a.Id, b.名称, b.性质, A.是否独立结算" & vbNewLine & _
'                "From 医疗卡类别 a, 结算方式 b" & vbNewLine & _
'                "Where a.结算方式 = b.名称 And a.名称 = [1]" & vbNewLine & _
'                "Union" & vbNewLine & _
'                "Select 2 As 类型, c.编号 As Id, d.名称, d.性质, 0 as 是否独立结算" & vbNewLine & _
'                "From 消费卡类别目录 c, 结算方式 d" & vbNewLine & _
'                "Where c.结算方式 = d.名称 And c.名称 = [1]"
'
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "GetPayInfo", str结算方式)
'        If rsTemp.EOF Then
'            MsgBox str结算方式 & "是无效的结算方式，请选择其他支付方式。", vbInformation, gstrSysName
'            Exit Function
'        End If
'        objPayInfo.名称 = str结算方式
'        objPayInfo.结算性质 = Val(Nvl(rsTemp!性质))
'        objPayInfo.结算方式 = Nvl(rsTemp!名称)
'        objPayInfo.接口序号 = Val(Nvl(rsTemp!ID))
'        objPayInfo.消费卡 = Val(Nvl(rsTemp!类型)) = 2
'        objPayInfo.独立结算 = Val(Nvl(rsTemp!是否独立结算)) = 1
'        If objPayInfo.接口序号 > 0 Then
'            If objPayInfo.消费卡 Then
'                objPayInfo.支付类型 = Pay_SquarePay
'            Else
'                objPayInfo.支付类型 = Pay_ThreePay
'            End If
'        Else
'            objPayInfo.支付类型 = Pay_CashPay
'        End If
    End If
    GetPayInfo = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function zlGetBackInvoice(ByVal strNO As String, ByRef strBackInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取回收发票
    ' 入参 : strNO-退号单据
    '        strBackInvoice-退号涉及的发票号
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/2/26 16:29
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    
    On Error GoTo errH
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
    zlGetBackInvoice = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBookingNO(ByVal strInput As String) As String
    Dim objInterCard As clsInterFaceCard
    Dim lng预约失约次数 As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strPatiIds As String
    
    If zlCreateOneCardObject(Nothing, glngSys, glngModul, gcnOracle, gstrDBUser, objInterCard) = False Then Exit Function
    lng预约失约次数 = Val(zlDatabase.GetPara("预约失约次数", glngSys, glngModul, 0))
    If Len(strInput) = 8 And InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid(strInput, 1, 1))) > 0 And IsNumeric(Mid(strInput, 2)) Then
        strInput = UCase(strInput)
        strSQL = "" & _
                "Select Min(A.NO) NO" & vbNewLine & _
                "From 门诊费用记录 A" & vbNewLine & _
                "Where A.记录性质 = 4 And A.记录状态 = 0 And A.No = [1] " & _
                IIf(lng预约失约次数 > 0, " And A.发生时间 between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
                "  And ((nvl(A.加班标志,0) =0 And A.发生时间 > Trunc(Sysdate) - [2]) or  (nvl(A.加班标志,0) =1 And A.发生时间 > Trunc(Sysdate) - [3])) ")
    Else
        If objInterCard.GetPatiIdsByRange(strInput, strPatiIds) = False Then Exit Function
        If strPatiIds = "" Then Exit Function
        strInput = strPatiIds
        strSQL = "" & _
            "Select /*+cardinality(B,10) */ Min(A.NO) NO" & vbNewLine & _
            "From 门诊费用记录 A, Table(f_num2list([1])) B" & vbNewLine & _
            "Where A.记录性质 = 4 And A.记录状态 = 0 And A.病人id = B.Column_Value(+) " & _
            IIf(lng预约失约次数 > 0, " And A.发生时间 between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
            "  And ((nvl(A.加班标志,0) =0 And A.发生时间 > Trunc(Sysdate) - [2]) or (nvl(A.加班标志,0) =1 And A.发生时间 > Trunc(Sysdate) - [3])) ")
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetBookingNO", strInput, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency)
    GetBookingNO = "" & rsTmp!NO
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAbNormalNO(ByVal strInput As String) As String
    Dim objInterCard As clsInterFaceCard
    Dim lng预约失约次数 As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strPatiIds As String
    
    If zlCreateOneCardObject(Nothing, glngSys, glngModul, gcnOracle, gstrDBUser, objInterCard) = False Then Exit Function
    lng预约失约次数 = Val(zlDatabase.GetPara("预约失约次数", glngSys, glngModul, 0))
    If Len(strInput) = 8 And InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid(strInput, 1, 1))) > 0 And IsNumeric(Mid(strInput, 2)) Then
        strInput = UCase(strInput)
        strSQL = " And A.NO = [1]"
        strSQL = "" & _
                "Select Min(A.NO) NO" & vbNewLine & _
                "From 门诊费用记录 A" & vbNewLine & _
                "Where A.记录性质 = 4 And A.记录状态 = 1 And A.No = [1] " & _
                IIf(lng预约失约次数 > 0, " And A.发生时间 between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
                "  And ((nvl(A.加班标志,0) =0 And A.发生时间 > Trunc(Sysdate) - [2]) or  (nvl(A.加班标志,0) =1 And A.发生时间 > Trunc(Sysdate) - [3])) ")
    Else
        If objInterCard.GetPatiIdsByRange(strInput, strPatiIds) = False Then Exit Function
        If strPatiIds = "" Then Exit Function
        strInput = strPatiIds
        strSQL = "" & _
            "Select /*+cardinality(B,10) */ Min(A.NO) NO" & vbNewLine & _
            "From 门诊费用记录 A, Table(f_num2list([1])) B" & vbNewLine & _
            "Where A.记录性质 = 4 And A.记录状态 = 1 And A.病人id = B.Column_Value(+) " & _
            IIf(lng预约失约次数 > 0, " And A.发生时间 between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
            "  And ((nvl(A.加班标志,0) =0 And A.发生时间 > Trunc(Sysdate) - [2]) or (nvl(A.加班标志,0) =1 And A.发生时间 > Trunc(Sysdate) - [3])) ")
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAbNormalNO", strInput, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency)
    GetAbNormalNO = "" & rsTmp!NO
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetBalanceInfor(ByVal bytMode As Byte, ByVal objRegInfor As clsRegEventInfor, _
                        ByVal strAdvance As String, ByRef str缺省退现 As String, _
                        str退现 As String, strBalance As String, str退消费卡 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取退号的结算信息
    ' 入参 : bytMode-退号模式：4：正常退费；5-作废；6-重退
    '        objRegInfor-本次挂号信息
    '        strAdvance- 医保不支持退现的结算方式
    ' 出参 : str退现 - 退现内容
    '        strBalance - 退接口支付,需要先保存校对数据
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/2/26 16:29
    '---------------------------------------------------------------------------------------
    Dim objPay As clsPayInfo, objSubPay As clsSubPayInfo
    Dim str现金 As String, dbl现金 As Double
    
    On Error GoTo errH
    For Each objPay In objRegInfor.objPayInfos
        If bytMode = 4 Or (objPay.校对标志 <> 2 And bytMode = 6) Or (bytMode = 5 And objPay.校对标志 <> 1) Then
            If objPay.支付类型 = Pay_CashPay Then
                If objPay.结算性质 = 1 Then
                    str现金 = objPay.结算方式: str缺省退现 = objPay.结算方式
                    dbl现金 = dbl现金 + objPay.支付金额
                ElseIf objPay.支付金额 <> 0 Then
                    str退现 = str退现 & "|" & objPay.结算方式 & "," & objPay.支付金额 & "," & objPay.结算号码 & ", "
                End If
            ElseIf objPay.支付类型 = Pay_InsurePay Then
                For Each objSubPay In objPay
                    If InStr(strAdvance, objSubPay.结算方式) <> 0 Then
                        dbl现金 = dbl现金 + objSubPay.结算金额
                        If str现金 = "" Then str现金 = str缺省退现
                    Else
                        strBalance = strBalance & "|" & objSubPay.结算方式 & "," & objSubPay.结算金额 & "," & 0
                    End If
                Next
            ElseIf objPay.支付类型 = Pay_SquarePay Then
                str退消费卡 = str退消费卡 & "|" & objPay.结算方式 & "," & objPay.支付金额
            ElseIf objPay.支付类型 = Pay_ThreePay Then
                For Each objSubPay In objPay
                    strBalance = strBalance & "|" & objSubPay.结算方式 & "," & objSubPay.结算金额 & "," & 1
                Next
            End If
        End If
    Next
    If RoundEx(dbl现金, 6) <> 0 Then
        str退现 = str退现 & "|" & str现金 & "," & dbl现金
    End If
    If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    If str退消费卡 <> "" Then str退消费卡 = Mid(str退消费卡, 2)
    If str退现 <> "" Then str退现 = Mid(str退现, 2)
    
    zlGetBalanceInfor = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetAll医生() As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    If mrsDoctor Is Nothing Then
        '人员性质固定，直接写在sql中，不以绑定变量传入
        strSQL = "Select a.Id, a.姓名, Upper(a.简码) As 简码,b.部门id,a.编号" & _
                " From 人员表 a, 部门人员 b, 人员性质说明 c" & _
                " Where a.Id = b.人员id And a.Id = c.人员id And c.人员性质 = '医生' And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order By a.简码 Desc"
        Set mrsDoctor = zlDatabase.OpenSQLRecord(strSQL, "GetAll医生")
    End If
    Set GetAll医生 = mrsDoctor
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckRegistAppointment(ByVal strNO As String) As Boolean
    '检查预约记录是否被接收
    'True-预约记录未接收;False-预约记录已被接收
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo Errhand
    strSQL = "Select 1 From 病人挂号记录 Where NO = [1] And 接收时间 Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckRegistAppointment", strNO)
    If Not rsTmp.EOF Then
        CheckRegistAppointment = True
    Else
        CheckRegistAppointment = False
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetColItem(colInfo As Collection, strItem As String) As String
    If colInfo Is Nothing Then Exit Function
    
    Err.Clear: On Error Resume Next
    GetColItem = colInfo("_" & strItem)
    Err.Clear: On Error GoTo 0
End Function

Public Function GetRoom(ByVal lng安排ID As Long, ByVal lng计划ID As Long) As String
'功能：根据号别的分诊方式获取号别的诊室
    Dim strSQL As String, str号别 As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
            
    '139670：2019/6/27，李南春，收费预约未来的挂号时，根据计划ID获取分诊方式以及指定分诊的诊室
    '其他分诊方式因为不能统计候诊人数以及统计的几家医院都没使用，本次不处理
    If lng计划ID = 0 Then
        strSQL = "Select 号码,Nvl(分诊方式,0) as 分诊 From 挂号安排 Where ID=[1]"
    Else
        strSQL = "Select 号码,Nvl(分诊方式,0) as 分诊 From 挂号安排计划 Where ID=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoom", lng安排ID, lng计划ID)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!分诊 = 0 Then Exit Function '不分诊
    
    str号别 = Nvl(rsTmp!号码)
    '处理分诊
    If rsTmp!分诊 = 1 Then
        '指定诊室
        If lng计划ID = 0 Then
            strSQL = "Select 门诊诊室 From 挂号安排诊室 Where 号表ID=[1]"
        Else
            strSQL = "Select 门诊诊室 From 挂号计划诊室 Where 计划ID=[2]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoom", lng安排ID, lng计划ID)
        If Not rsTmp.EOF Then GetRoom = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 2 Then
        '动态分诊：该个号别当天挂号未诊数最少的诊室   //todo未考虑预约挂号
        strSQL = _
            " Select 门诊诊室,Sum(NUM) as NUM From (" & _
                " Select 门诊诊室,0 as NUM From 挂号安排诊室 Where 号表ID=[1]" & _
                " Union ALL" & _
                " Select 诊室,Count(诊室) as NUM From 病人挂号记录" & _
                " Where Nvl(执行状态,0)=0 And 记录性质=1 and 记录状态=1 and  发生时间 Between Trunc(Sysdate) And Sysdate And 号别=[2]" & _
                " And 诊室 IN(Select 门诊诊室 From 挂号安排诊室 Where 号表ID=[1])" & _
                " Group by 诊室)" & _
            " Group by 门诊诊室 Order by Num"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoom", lng安排ID, str号别)
        If Not rsTmp.EOF Then GetRoom = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 3 Then
        '平均分诊：当前分配=1表示下次应取的当前诊室
        strSQL = "Select 号表ID,门诊诊室,当前分配 From 挂号安排诊室 Where 号表ID=" & lng安排ID
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "GetRoom", adOpenDynamic, adLockOptimistic)
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!当前分配), 0, rsTmp!当前分配) = 1 Then
                    GetRoom = rsTmp!门诊诊室
                    rsTmp!当前分配 = 0
                    
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!当前分配 = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '处理第一次平均分配
            If GetRoom = "" Then
                rsTmp.MoveFirst
                GetRoom = rsTmp!门诊诊室
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!当前分配 = 1
                rsTmp.Update
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRoomVisit(ByVal lng记录ID As Long, ByVal bln预约接收立即就诊 As Boolean) As String
'功能：根据号别的分诊方式获取号别的诊室
    Dim strSQL As String, strRoomIDs As String
    Dim rsTmp As ADODB.Recordset, rsRoom As ADODB.Recordset
    
    On Error GoTo errH
    
    If bln预约接收立即就诊 Then
        strSQL = "Select a.Id" & vbNewLine & _
                "From 临床出诊记录 a, 临床出诊记录 b" & vbNewLine & _
                "Where a.号源id = b.号源id And a.是否分时段 = b.是否分时段 And a.是否序号控制 = b.是否序号控制 And a.科室id = b.科室id And" & vbNewLine & _
                "      Nvl(a.医生id, 0) = Nvl(b.医生id, 0) And a.上班时段 = b.上班时段 And Nvl(a.是否发布, 0) = 1 And a.出诊日期 = Trunc(Sysdate) And" & vbNewLine & _
                "      b.Id = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "更换出诊ID", lng记录ID)
        If Not rsTmp.EOF Then
            lng记录ID = Val(Nvl(rsTmp!id))
        End If
    End If
            
    strSQL = "Select ID,Nvl(分诊方式,0) as 分诊 From 临床出诊记录 Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoomVisit", lng记录ID)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!分诊 = 0 Then Exit Function '不分诊
    
    '处理分诊
    If rsTmp!分诊 = 1 Then
        '指定诊室
        strSQL = "Select B.名称 As 门诊诊室 From 临床出诊诊室记录 A,门诊诊室 B Where A.诊室ID=B.ID And A.记录ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoomVisit", CLng(rsTmp!id))
        If Not rsTmp.EOF Then GetRoomVisit = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 2 Then
        '动态分诊：该个号别当天挂号未诊数最少的诊室   //todo未考虑预约挂号
        strSQL = _
            " Select 门诊诊室,Sum(NUM) as NUM From (" & _
                " Select B.名称 As 门诊诊室,0 as NUM From 临床出诊诊室记录 A,门诊诊室 B Where A.诊室ID = B.ID And 记录ID=[1]" & _
                " Union ALL" & _
                " Select 诊室,Count(诊室) as NUM From 病人挂号记录" & _
                " Where Nvl(执行状态,0)=0 And 记录性质=1 and 记录状态=1 and  发生时间 Between Trunc(Sysdate) And Sysdate And 出诊记录ID = [2]" & _
                " And 诊室 IN (Select D.名称 As 门诊诊室 From 临床出诊诊室记录 C,门诊诊室 D Where C.记录ID=[1] And C.诊室ID = D.ID )" & _
                " Group by 诊室)" & _
            " Group by 门诊诊室 Order by Num"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoomVisit", CLng(rsTmp!id), lng记录ID)
        If Not rsTmp.EOF Then GetRoomVisit = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 3 Then
        '平均分诊：当前分配=1表示下次应取的当前诊室
        strSQL = "Select * From 临床出诊诊室记录 Where 记录ID=" & rsTmp!id
'        strSQL = "Select A.记录ID,B.名称 As 门诊诊室,A.当前分配 From 临床出诊诊室记录 A,门诊诊室 B Where A.诊室ID=B.ID And A.记录ID=" & rsTmp!ID
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "GetRoomVisit", adOpenDynamic, adLockOptimistic)
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!当前分配), 0, rsTmp!当前分配) = 1 Then
                    strRoomIDs = rsTmp!诊室ID
                    rsTmp!当前分配 = 0
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!当前分配 = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '处理第一次平均分配
            If strRoomIDs = "" Then
                rsTmp.MoveFirst
                strRoomIDs = rsTmp!诊室ID
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!当前分配 = 1
                rsTmp.Update
            End If
        End If
        If strRoomIDs <> "" Then
            strSQL = "Select 名称 From 门诊诊室 Where ID = [1]"
            Set rsRoom = zlDatabase.OpenSQLRecord(strSQL, "GetRoomVisit", strRoomIDs)
            If Not rsRoom.EOF Then
                GetRoomVisit = rsRoom!名称
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckCanModifyName(ByVal strNO As String, ByVal str就诊时间 As String, ByVal lng病人ID As Long) As Boolean
'功能:检查挂号单是否可以修改姓名,如果不是挂号时建的档,就不能修改.
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1" & vbNewLine & _
            "From 门诊费用记录 A" & vbNewLine & _
            "Where A.NO = [1] And A.记录性质 = 4 And A.登记时间 = To_Date([2],'YYYY-MM-DD HH24:MI:SS') And A.病人id = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "挂号时建档", strNO, str就诊时间, lng病人ID)
    CheckCanModifyName = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGet失约号(ByVal bytRegMode As Byte, ByVal varUDID As Variant, ByVal lng预约有效时间 As Long, datThis As Date) As Long
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取安排在某一天.预约失约数
    ' 入参 : bytRegMode:0-计划安排模式；1-出诊排班模式
    '        var标识:挂号唯一标识：bytRegMode=0是号别；bytRegMode=1 是记录id
    '        datThis-指定日期
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/16 17:17
    '---------------------------------------------------------------------------------------
    Dim strSQL  As String, strWhere As String
    Dim rsTmp   As ADODB.Recordset
    Dim strBegin  As String, strEnd As String
    
    If bytRegMode = 0 Then
        strWhere = "号别 = [1]"
    Else
        strWhere = "出诊记录ID = [1]"
    End If
    
    strSQL = "Select Count(1) As 失约号" & vbNewLine & _
            " From 病人挂号记录" & vbNewLine & _
            " Where " & strWhere & " And 记录性质 = 2 And 记录状态 = 1 And 发生时间 - [2] / 24 / 60 < Sysdate And 发生时间 Between to_Date([3],'YYYY-MM-DD') And to_Date([4],'YYYY-MM-DD') - 1/24/60/60"
    strBegin = Format(datThis, "yyyy-MM-dd")
    strEnd = Format(datThis + 1, "yyyy-MM-dd")
    On Error GoTo Hd
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "失约数量", varUDID, lng预约有效时间, strBegin, strEnd)
    If rsTmp.EOF Then
        zlGet失约号 = 0
        Set rsTmp = Nothing
        Exit Function
    End If
    zlGet失约号 = Val(Nvl(rsTmp!失约号, 0))
    Set rsTmp = Nothing
   Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


Public Function zlCheckRegistVaild(ByVal objService As zlPublicExpense.clsService, _
                ByVal byt操作类型 As Byte, ByVal lng病人ID As Long, ByVal str号码 As String, _
                ByVal dat预约时间 As Date, ByVal bln专家号 As Boolean, _
                Optional ByVal str出诊记录ID As String, Optional ByVal str预约方式 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : Zl_Fun_病人挂号记录_Check
    ' 入参 : byt操作类型：0-挂号;1-预约;2-预约接收
    '        str出诊记录ID：老排班模式不传，新排班>=0
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/26 18:43
    '---------------------------------------------------------------------------------------
    Dim rsCheck As ADODB.Recordset, strSQL As String
    Dim strResult As String
    On Error GoTo Errhand

    If objService Is Nothing Then
        Call ShowMsgbox("服务接口(zlPublicExpense.clsService)未启用！")
        Exit Function
    End If
    If byt操作类型 = 1 Then
        If objService.zlPatisvr_GetPatiBlackInfo(lng病人ID, "预约", str预约方式) = False Then Exit Function
    End If
    
    strSQL = "Select Zl_Fun_病人挂号记录_Check_S([1],[2],[3],[4],[5],[6]) As 检查结果 From Dual"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "zlCheckRegistVaild", byt操作类型, lng病人ID, str号码, str出诊记录ID, dat预约时间, IIf(bln专家号, 1, 0))
    If Not rsCheck.EOF Then
        strResult = Nvl(rsCheck!检查结果)
        If Val(Mid(strResult, 1, 1)) <> 0 Then
            MsgBox Mid(strResult, 3), vbInformation, gstrSysName
            Exit Function
        End If
    Else
        MsgBox "有效性检查失败,无法继续！", vbInformation, gstrSysName
        Exit Function
    End If
    zlCheckRegistVaild = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Check复诊(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long) As Boolean
'功能:判断病人是否再次到“相同临床性质的临床科室”挂号
'     包括挂过号的,或住过院的,复诊不好确定时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo Errhand
    strSQL = "Select Zl1_Fun_Getreturnvisit([1],[2]) As 复诊标志 From Dual"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "复诊检查", lng病人ID, lng执行部门ID)
    Check复诊 = Val(Nvl(rsTmp!复诊标志)) = 1
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlCisNewVisitRec(ByVal objService As zlPublicExpense.clsService, ByVal strNO As String, _
                Optional ByRef str就诊时间 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 挂号完成后通知临床
    ' 入参 :
    ' 出参 : str就诊时间-发生时间
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo Errhand

    strSQL = "Select 记录性质,号别,号类,预约,病人id,门诊号,姓名,性别,年龄,费别,复诊,急诊,诊室,执行部门ID," & _
                    "执行人,发生时间,结算模式,出诊记录ID " & vbNewLine & _
            "From 病人挂号记录 Where No  = [1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "更新就诊登记", strNO)
    If rsTmp.EOF Then Exit Function
    Set cllVisit = New Collection
    With rsTmp
        cllVisit.Add strNO, "挂号单号"
        cllVisit.Add Val(Nvl(!记录性质)), "单据性质"
        cllVisit.Add Val(Nvl(!预约)), "预约标志"
        cllVisit.Add Nvl(!号别), "号别"
        cllVisit.Add Val(Nvl(!出诊记录ID)), "出诊记录ID"
        cllVisit.Add Nvl(!号类), "号类"
        cllVisit.Add Val(Nvl(!病人ID)), "病人ID"
        cllVisit.Add Nvl(!门诊号), "门诊号"
        cllVisit.Add Nvl(!姓名), "姓名"
        cllVisit.Add Nvl(!性别), "性别"
        cllVisit.Add Nvl(!年龄), "年龄"
        cllVisit.Add Val(Nvl(!复诊)), "复诊"
        cllVisit.Add Nvl(!费别), "费别"
        cllVisit.Add Val(Nvl(!急诊)), "急诊"
        cllVisit.Add Nvl(!诊室), "诊室"
        cllVisit.Add Val(Nvl(!执行部门ID)), "执行部门ID"
        cllVisit.Add Nvl(!执行人), "执行人"
        cllVisit.Add Nvl(!发生时间), "发生时间"
        cllVisit.Add Val(Nvl(!结算模式)), "结算模式"
        str就诊时间 = Nvl(!发生时间)
    End With
    zlCisNewVisitRec = objService.zlCISSvr_NewOutPatiVisitRec(cllVisit)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCisUpdateVisitRoom(ByVal objService As zlPublicExpense.clsService, _
                ByVal strNO As String, ByVal str执行人 As String, ByVal strRoom As String, _
                Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 更新临床就诊诊室
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    On Error GoTo Errhand

    Set cllVisit = New Collection
    cllVisit.Add Array("挂号单号", strNO)
    cllVisit.Add Array("诊室", strRoom)
    cllVisit.Add Array("执行人", str执行人)
    
    zlCisUpdateVisitRoom = objService.zlCISSvr_UpdateOutPatiVisitRec(cllVisit, True, strErrMsg)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCisUpdateVisitState(ByVal objService As zlPublicExpense.clsService, _
                ByVal strNO As String, ByVal int执行状态 As Integer, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 更新临床就诊状态
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    On Error GoTo Errhand

    Set cllVisit = New Collection
    cllVisit.Add Array("挂号单号", strNO)
    cllVisit.Add Array("执行状态", int执行状态)
    
    zlCisUpdateVisitState = objService.zlCISSvr_UpdateOutPatiVisitRec(cllVisit, True, strErrMsg)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCisUpdateVistPatiBase(ByVal objService As zlPublicExpense.clsService, _
                ByVal blnNewPati As Boolean, ByVal strNO As String, ByVal lng病人ID As Long, ByVal str门诊号 As String, _
                ByVal str姓名 As String, ByVal str性别 As String, ByVal str年龄 As String, _
                ByVal str费别 As String, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 更新临床就诊信息中病人记录信息
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    On Error GoTo Errhand

    Set cllVisit = New Collection
    cllVisit.Add Array("挂号单号", strNO)
    If blnNewPati Then
        cllVisit.Add Array("病人ID", lng病人ID)
        cllVisit.Add Array("姓名", str姓名)
        cllVisit.Add Array("性别", str性别)
        cllVisit.Add Array("年龄", str年龄)
    End If
    cllVisit.Add Array("门诊号", str门诊号)
    cllVisit.Add Array("费别", str费别)
    
    zlCisUpdateVistPatiBase = objService.zlCISSvr_UpdateOutPatiVisitRec(cllVisit, True, strErrMsg)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCisDoneVisit(ByVal objService As zlPublicExpense.clsService, _
                ByVal strNO As String, ByVal str执行人 As String, ByVal strRoom As String, ByVal str完成时间 As String, _
                ByVal str摘要 As String, ByVal bln护士执行 As Boolean, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 完成就诊
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    On Error GoTo Errhand
    
    Set cllVisit = New Collection
    cllVisit.Add Array("挂号单号", strNO)
    cllVisit.Add Array("执行状态", 1)
    cllVisit.Add Array("诊室", strRoom)
    cllVisit.Add Array("完成时间", str完成时间)
    If str执行人 <> "" Then
        cllVisit.Add Array("执行人", str执行人)
    End If
    If str摘要 <> "" Then
        cllVisit.Add Array("摘要", str摘要)
    End If
    
    zlCisDoneVisit = objService.zlCISSvr_UpdateOutPatiVisitRec(cllVisit, True, strErrMsg)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCisCancelVisit(ByVal objService As zlPublicExpense.clsService, _
                ByVal strNO As String, ByVal bln护士执行 As Boolean, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 取消接诊
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    On Error GoTo Errhand
    
    Set cllVisit = New Collection
    cllVisit.Add Array("挂号单号", strNO)
    cllVisit.Add Array("执行状态", IIf(bln护士执行, 0, 2))
    cllVisit.Add Array("完成时间", "")
    
    zlCisCancelVisit = objService.zlCISSvr_UpdateOutPatiVisitRec(cllVisit, True, strErrMsg)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AddAdressInfo(ByRef cllPati As Collection, padd现住址 As Object, padd户口地址 As Object) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取结构化地址信息
    ' 入参 : padd现住址:现地址控件
    '        padd户口地址:户口地址控件
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/30 21:59
    '---------------------------------------------------------------------------------------
    Dim cllTemp As Collection, cllSubTemp As Collection
    On Error GoTo Errhand

    If cllPati Is Nothing Then Set cllPati = New Collection
    Set cllTemp = New Collection
    
    If padd现住址.Value <> "" Then
        If zlGetAdressCol(cllSubTemp, 1, 3, padd现住址.value省, padd现住址.value市, padd现住址.value区县, padd现住址.value乡镇, _
            padd现住址.value详细地址, padd现住址.Code) = False Then Exit Function
    Else
        If zlGetAdressCol(cllSubTemp, 2, 3) = False Then Exit Function
    End If
    cllTemp.Add cllSubTemp, "家庭地址"
    
    If padd户口地址.Value <> "" Then
        If zlGetAdressCol(cllSubTemp, 1, 4, padd户口地址.value省, padd户口地址.value市, padd户口地址.value区县, padd户口地址.value乡镇, _
            padd户口地址.value详细地址, padd户口地址.Code) = False Then Exit Function
    Else
        If zlGetAdressCol(cllSubTemp, 2, 4) = False Then Exit Function
    End If
    cllTemp.Add cllSubTemp, "户口地址"
    
    If cllTemp.Count > 0 Then cllPati.Add cllTemp, "地址信息"
    AddAdressInfo = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiCol(ByRef cllPati_Out As Collection, _
                ByVal lng病人ID As Long, ByVal str姓名 As String, ByVal str性别 As String, _
                ByVal str年龄 As String, ByVal str出生日期 As String, ByVal str身份证号 As String, ByVal str门诊号 As String, _
                ByVal str费别 As String, ByVal str医疗付款方式名称 As String, ByVal str国籍 As String, _
                ByVal str民族 As String, ByVal str婚姻状况 As String, ByVal str职业 As String, _
                ByVal str身份 As String, ByVal str工作单位 As String, ByVal str单位电话 As String, _
                ByVal lng合同单位id As Long, ByVal str单位邮编 As String, _
                ByVal str家庭地址 As String, ByVal str家庭电话 As String, _
                ByVal str家庭地址邮编 As String, ByVal str区域 As String, ByVal str出生地点 As String, _
                ByVal str户口地址 As String, ByVal str户口地址邮编 As String, ByVal str联系人姓名 As String, _
                ByVal str联系人身份证号 As String, ByVal str联系人电话 As String, ByVal str联系人关系 As String, _
                ByVal str监护人 As String, ByVal str手机号 As String, ByVal str医保号 As String, ByVal int险类 As Integer, _
                ByVal str登记时间 As String, Optional ByVal lng社区序号 As Long, Optional ByVal str社区号码 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人信息集合
    ' 入参 : 病人基本信息
    ' 出参 : 病人信息集合
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/30 19:04
    '---------------------------------------------------------------------------------------
    Dim cllTmp As Collection
    On Error GoTo Errhand
    Set cllPati_Out = New Collection
    
    cllPati_Out.Add Array("病人ID", lng病人ID), "病人ID"
    cllPati_Out.Add Array("姓名", str姓名), "姓名"
    cllPati_Out.Add Array("性别", str性别), "性别"
    cllPati_Out.Add Array("年龄", str年龄), "年龄"
    cllPati_Out.Add Array("出生日期", str出生日期), "出生日期"
    cllPati_Out.Add Array("身份证号", str身份证号), "身份证号"
    cllPati_Out.Add Array("门诊号", str门诊号), "门诊号"
    cllPati_Out.Add Array("费别", str费别), "费别"
    cllPati_Out.Add Array("医疗付款方式名称", str医疗付款方式名称), "医疗付款方式名称"
    cllPati_Out.Add Array("国籍", str国籍), "国籍"
    cllPati_Out.Add Array("民族", str民族), "民族"
    cllPati_Out.Add Array("婚姻状况", str婚姻状况), "婚姻状况"
    cllPati_Out.Add Array("职业", str职业), "职业"
    cllPati_Out.Add Array("身份", str身份), "身份"
    cllPati_Out.Add Array("工作单位", str工作单位), "工作单位"
    cllPati_Out.Add Array("单位电话", str单位电话), "单位电话"
    cllPati_Out.Add Array("合同单位ID", lng合同单位id), "合同单位ID"
    cllPati_Out.Add Array("单位邮编", str单位邮编), "单位邮编"
    cllPati_Out.Add Array("家庭地址", str家庭地址), "家庭地址"
    cllPati_Out.Add Array("家庭电话", str家庭电话), "家庭电话"
    cllPati_Out.Add Array("家庭地址邮编", str家庭地址邮编), "家庭地址邮编"
    cllPati_Out.Add Array("区域", str区域), "区域"
    cllPati_Out.Add Array("出生地点", str出生地点), "出生地点"
    cllPati_Out.Add Array("户口地址", str户口地址), "户口地址"
    cllPati_Out.Add Array("户口地址邮编", str户口地址邮编), "户口地址邮编"
    cllPati_Out.Add Array("监护人", str监护人), "监护人"
    cllPati_Out.Add Array("手机号", str手机号), "手机号"
    cllPati_Out.Add Array("医保号", str医保号), "医保号"
    cllPati_Out.Add Array("险类", int险类), "险类"
    cllPati_Out.Add Array("登记时间", str登记时间), "登记时间"
    cllPati_Out.Add Array("操作员姓名", UserInfo.姓名), "操作员姓名"
    cllPati_Out.Add Array("操作员编号", UserInfo.编号), "操作员编号"
    If str联系人姓名 <> "" Then
        Set cllTmp = New Collection
        cllTmp.Add str联系人姓名, "联系人姓名"
        cllTmp.Add str联系人身份证号, "联系人身份证号"
        cllTmp.Add str联系人电话, "联系人电话"
        cllTmp.Add str联系人关系, "联系人关系"
        cllPati_Out.Add cllTmp, "联系人"
    End If
    
    If str社区号码 <> "" Then
        Set cllTmp = New Collection
        cllTmp.Add lng社区序号, "社区序号"
        cllTmp.Add str社区号码, "社区号码"
        cllTmp.Add 1, "社区操作类型"
        cllPati_Out.Add cllTmp, "社区信息"
    End If
    
    zlGetPatiCol = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiMecCol(ByRef cllPati_Out As Collection, _
                ByVal lng病人ID As Long, ByVal str姓名 As String, ByVal str性别 As String, _
                ByVal str年龄 As String, ByVal str出生日期 As String, ByVal str身份证号 As String, _
                ByVal str门诊号 As String, ByVal str费别 As String, ByVal str医疗付款方式名称 As String, _
                ByVal str国籍 As String, ByVal str民族 As String, ByVal str婚姻状况 As String, _
                ByVal str职业 As String, ByVal str身份 As String, _
                ByVal str工作单位 As String, ByVal str单位电话 As String, _
                ByVal lng合同单位id As Long, ByVal str单位邮编 As String, _
                ByVal str家庭地址 As String, ByVal str家庭电话 As String, ByVal str家庭地址邮编 As String, _
                ByVal str户口地址 As String, ByVal str户口地址邮编 As String, _
                ByVal str监护人 As String, ByVal str手机号 As String, ByVal str医保号 As String, ByVal int险类 As Integer, _
                ByVal str登记时间 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人信息集合
    ' 入参 : 病人基本信息
    ' 出参 : 病人信息集合
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/30 19:04
    '---------------------------------------------------------------------------------------
    Dim cllTmp As Collection
    On Error GoTo Errhand
    Set cllPati_Out = New Collection
    
    cllPati_Out.Add Array("病人ID", lng病人ID)
    cllPati_Out.Add Array("姓名", str姓名)
    cllPati_Out.Add Array("性别", str性别)
    cllPati_Out.Add Array("年龄", str年龄)
    cllPati_Out.Add Array("出生日期", str出生日期)
    cllPati_Out.Add Array("身份证号", str身份证号)
    cllPati_Out.Add Array("门诊号", str门诊号)
    cllPati_Out.Add Array("费别", str费别)
    cllPati_Out.Add Array("医疗付款方式名称", str医疗付款方式名称)
    cllPati_Out.Add Array("国籍", str国籍)
    cllPati_Out.Add Array("民族", str民族)
    cllPati_Out.Add Array("婚姻状况", str婚姻状况)
    cllPati_Out.Add Array("职业", str职业)
    cllPati_Out.Add Array("身份", str身份)
    cllPati_Out.Add Array("工作单位", str工作单位)
    cllPati_Out.Add Array("单位电话", str单位电话)
    cllPati_Out.Add Array("合同单位ID", lng合同单位id)
    cllPati_Out.Add Array("单位邮编", str单位邮编)
    cllPati_Out.Add Array("家庭地址", str家庭地址)
    cllPati_Out.Add Array("家庭电话", str家庭电话)
    cllPati_Out.Add Array("家庭地址邮编", str家庭地址邮编)
    cllPati_Out.Add Array("户口地址", str户口地址)
    cllPati_Out.Add Array("户口地址邮编", str户口地址邮编)
    cllPati_Out.Add Array("监护人", str监护人)
    cllPati_Out.Add Array("手机号", str手机号)
    cllPati_Out.Add Array("医保号", str医保号)
    cllPati_Out.Add Array("险类", int险类)
    cllPati_Out.Add Array("登记时间", str登记时间)
    cllPati_Out.Add Array("操作员姓名", UserInfo.姓名)
    cllPati_Out.Add Array("操作员编号", UserInfo.编号)
    
    zlGetPatiMecCol = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetAdressCol(ByRef cllAdress_Out As Collection, ByVal byt操作功能 As Byte, _
                ByVal byt地址类别 As Byte, Optional ByVal str地址_省 As String, Optional ByVal str地址_市 As String, _
                Optional ByVal str地址_县 As String, Optional ByVal str地址_乡 As String, Optional ByVal str地址_其他 As String, _
                Optional ByVal str区划代码 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取结构化地址信息集合
    ' 入参 : 结构化地址信息
    ' 出参 : 结构化地址集合
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/30 19:04
    '---------------------------------------------------------------------------------------
    On Error GoTo Errhand
    Set cllAdress_Out = New Collection
    
    cllAdress_Out.Add byt操作功能, "操作功能"
    cllAdress_Out.Add byt地址类别, "地址类别"
    cllAdress_Out.Add str地址_省, "地址_省"
    cllAdress_Out.Add str地址_市, "地址_市"
    cllAdress_Out.Add str地址_县, "地址_县"
    cllAdress_Out.Add str地址_乡, "地址_乡"
    cllAdress_Out.Add str地址_其他, "地址_其他"
    cllAdress_Out.Add str区划代码, "区划代码"

    zlGetAdressCol = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet缺省费别() As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取缺省费别
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/31 14:11
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo Errhand
    strSQL = "Select 名称  From 费别 Where 缺省标志 = 1 And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "缺省费别")
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
     End Select
End Function

Public Function zlExcMedcCardService(ByVal objService As zlPublicExpense.clsService, _
                ByVal int操作类型 As Integer, ByVal lng病人ID As Long, ByVal lng卡类别id As Long, _
                ByVal str原卡号 As String, ByVal str医疗卡号 As String, ByVal str密码 As String, _
                ByVal str变动原因 As String, Optional ByVal strIC卡号 As String, Optional ByVal str挂失方式 As String, _
                Optional ByVal str二维码 As String, Optional ByVal str终止使用时间 As String, _
                Optional ByVal strNO As String, Optional ByVal dbl卡费 As Double, _
                Optional ByVal str操作时间 As String, Optional ByRef strErrMsg As String) As Boolean
    Dim cllCard As Collection, cllVisit As Collection

    On Error GoTo errHandle
    If objService Is Nothing Then Exit Function
    str密码 = zlCommFun.zlStringEncode(str密码)
    
    Set cllCard = New Collection
    cllCard.Add Array("操作类型", int操作类型)
    cllCard.Add Array("病人id", lng病人ID)
    cllCard.Add Array("卡类别ID", lng卡类别id)
    cllCard.Add Array("原卡号", str原卡号)
    cllCard.Add Array("医疗卡号", str医疗卡号)
    cllCard.Add Array("变动原因", str变动原因)
    cllCard.Add Array("密码", str密码)
    cllCard.Add Array("IC卡号", strIC卡号)
    cllCard.Add Array("挂失方式", str挂失方式)
    cllCard.Add Array("二维码", str二维码)
    cllCard.Add Array("终止使用时间", str终止使用时间)
    cllCard.Add Array("操作时间", str操作时间)
    cllCard.Add Array("操作员姓名", UserInfo.姓名)
    cllCard.Add Array("操作员编号", UserInfo.编号)
    cllCard.Add Array("单据号", strNO)
    cllCard.Add Array("卡费", dbl卡费)

    If objService.zlPatisvr_SaveMedcCard(cllCard, strErrMsg) = False Then Exit Function
    zlExcMedcCardService = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function

Public Function zlPatiUpdVisitService(ByVal objService As zlPublicExpense.clsService, ByVal lng病人ID As Long, _
                ByVal strNode As String, Optional ByVal int就诊状态 As Integer, Optional ByVal str就诊时间 As String, _
                Optional ByVal str门诊号 As String, Optional ByVal str就诊诊室 As String, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 更新病人就诊信息
    ' 入参 : strNode-需要更新的节点
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/5 15:40
    '---------------------------------------------------------------------------------------
    Dim cllPati As Collection
    
    On Error GoTo Errhand
    If objService Is Nothing Then Exit Function
    Set cllPati = New Collection
    cllPati.Add lng病人ID, "病人ID"
    If InStr("," & strNode & ",", "就诊状态") > 0 Then
        cllPati.Add int就诊状态, "就诊状态"
    End If
    If InStr("," & strNode & ",", "门诊号") > 0 Then
        cllPati.Add str门诊号, "门诊号"
    End If
    If InStr("," & strNode & ",", "诊室") > 0 Then
        cllPati.Add str就诊诊室, "就诊诊室"
    End If
    If InStr("," & strNode & ",", "就诊时间") > 0 Then
        cllPati.Add str就诊时间, "就诊时间"
    End If
    
    If objService.ZlPatiSvr_UpdateOutPatiState(cllPati, strErrMsg) = False Then Exit Function
    
    zlPatiUpdVisitService = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function LoadErrUnBalanceInfo(ByRef rsBillAdvance As ADODB.Recordset, _
                objInterCard As clsInterFaceCard, ByVal objPayInfor As clsPayInfos, _
                ByVal str原结帐IDs As String, ByVal str销账IDs As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取异常重退未退数据
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/9 10:23
    '---------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objCard As Card
    On Error GoTo Errhand
    If objInterCard Is Nothing Then Exit Function
    
    strSQL = "Select Mod(a.记录性质, 10) As 记录性质, b.性质,Nvl(d.名称, a.结算方式) as 名称, a.结算方式, Sum(a.冲预交) As 金额, a.结算号码, Nvl(a.卡类别id, a.结算卡序号) As 卡类别id," & vbNewLine & _
            "       Decode(a.结算卡序号, Null, 0, 1) As 消费卡, a.卡号, e.消费卡id, a.关联交易id, a.交易流水号, a.交易说明" & vbNewLine & _
            "From 病人预交记录 a, 结算方式 b, 消费卡类别目录 d, 病人卡结算记录 e" & vbNewLine & _
            "Where a.结帐id In (Select /* +cardinality(M,10) */" & vbNewLine & _
            "                  m.Column_Value" & vbNewLine & _
            "                 From Table(f_Str2list([1])) m) And a.结算方式 Is Not Null And Nvl(a.校对标志, 0) <> 1 And a.结算方式 = b.名称(+) And" & vbNewLine & _
            "      a.Id = e.结算id(+) And a.结算卡序号 = d.编号(+) And a.结算方式 = d.结算方式(+) And " & vbNewLine & _
            "      (a.记录性质 = 11 Or (Nvl(a.卡类别id, Nvl(a.结算卡序号, 0)) = 0 And b.性质 In (1, 2)) Or" & vbNewLine & _
            "      (Nvl(a.卡类别id, Nvl(a.结算卡序号, 0)) > 0 And" & vbNewLine & _
            "      (Nvl(a.卡类别id, 0), Nvl(a.结算卡序号, 0)) Not In" & vbNewLine & _
            "      (Select Nvl(卡类别id, 0), Nvl(结算卡序号, 0)" & vbNewLine & _
            "         From 病人预交记录" & vbNewLine & _
            "         Where 结帐id In (Select /* +cardinality(M,10) */" & vbNewLine & _
            "                         m.Column_Value" & vbNewLine & _
            "                        From Table(f_Str2list([2])) m) And 记录状态 = 2 And 结算方式 Is Not Null)))" & vbNewLine & _
            "Group By a.记录性质, b.性质, a.结算方式,d.名称, a.结算号码, Nvl(a.卡类别id, a.结算卡序号), Decode(a.结算卡序号, Null, 0, 1), a.卡号, e.消费卡id, a.关联交易id," & vbNewLine & _
            "         a.交易流水号, a.交易说明"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取未退数据", str原结帐IDs, str销账IDs)
    
    If rsBillAdvance Is Nothing Then
        If InitBalanceData(rsBillAdvance) = False Then Exit Function
    End If
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            rsBillAdvance.AddNew
            rsBillAdvance!记录性质 = Val(Nvl(rsTmp!记录性质))
            rsBillAdvance!结算性质 = Val(Nvl(rsTmp!性质))
            rsBillAdvance!卡类别ID = Val(Nvl(rsTmp!卡类别ID))
            rsBillAdvance!消费卡 = Val(Nvl(rsTmp!消费卡))
            If Val(Nvl(rsTmp!卡类别ID)) > 0 And Val(Nvl(rsTmp!消费卡)) = 0 Then
                If objInterCard.GetCard(Val(Nvl(rsTmp!卡类别ID)), False, objCard) Then
                    If Nvl(rsBillAdvance!结算方式) = objCard.结算方式 Then
                        rsBillAdvance!支付方式 = objCard.名称
                    Else
                        rsBillAdvance!支付方式 = Nvl(rsTmp!名称)
                    End If
                Else
                    rsBillAdvance!支付方式 = Nvl(rsTmp!名称)
                End If
            Else
                 rsBillAdvance!支付方式 = Nvl(rsTmp!名称)
            End If
            rsBillAdvance!结算方式 = Nvl(rsTmp!结算方式)
            rsBillAdvance!金额 = Val(Nvl(rsTmp!金额))
            rsBillAdvance!结帐ID = IIf(Val(Nvl(rsTmp!记录性质)) = 5, objPayInfor.Card_结帐ID, objPayInfor.Reg_结帐ID)
            rsBillAdvance!消费卡ID = Val(Nvl(rsTmp!消费卡ID))
            rsBillAdvance!结算号码 = Nvl(rsTmp!结算号码)
            rsBillAdvance!卡号 = Nvl(rsTmp!卡号)
            rsBillAdvance!关联交易ID = Val(Nvl(rsTmp!关联交易ID))
            rsBillAdvance!交易流水号 = Nvl(rsTmp!交易流水号)
            rsBillAdvance!交易说明 = Nvl(rsTmp!交易说明)
            rsBillAdvance!校对标志 = 0
            rsBillAdvance!固定 = 0
            rsBillAdvance.Update
            rsTmp.MoveNext
        Loop
    End If
    LoadErrUnBalanceInfo = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

    
Public Function zlExcStationErrReceive(ByVal objService As zlPublicExpense.clsService, ByVal objExseSvr As zlPublicExpense.clsExpenceSvr, _
                ByVal strNO As String, ByRef lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 处理医生站预约接收异常数据
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/1/3 16:56
    '---------------------------------------------------------------------------------------
    Dim cllPro As Collection, cllSwapOther As Collection
    Dim blnTrans As Boolean
    Dim int同步状态 As Integer
    Dim str交易信息 As String, strErrMsg As String
    
    On Error GoTo Errhand
    lng病人ID = 0
    If objService.zlCISSvr_GetErrBillInfo(strNO, int同步状态, lng病人ID, , cllPro, cllSwapOther) = False Then Exit Function
    If int同步状态 = 0 Then zlExcStationErrReceive = True: Exit Function
    If int同步状态 <> 2 Then
        If int同步状态 <> -1 Then lng病人ID = 0
        If objService.zlCISSvr_delErrBillInfo(strNO, True) = False Then Exit Function
    Else
        gcnOracle.BeginTrans: blnTrans = True
            zlExecuteProcedureArrAy cllPro, "zlExcStationErrReceive", True, True
            If objService.zlCISSvr_delErrBillInfo(strNO, False, strErrMsg) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, gstrSysName
                Exit Function
            End If
            On Error Resume Next
            zlExecuteProcedureArrAy cllSwapOther, "zlExcStationErrReceive", True, True
            On Error GoTo Errhand
        gcnOracle.CommitTrans: blnTrans = False
    End If
    zlExcStationErrReceive = True
    Exit Function
Errhand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlInitPati(ByVal objInterCard As clsInterFaceCard, ByRef rsPatiInfor As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始病人信息集
    '返回:病人信息集
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsPatiInfor = New ADODB.Recordset
    If objInterCard Is Nothing Then
        Set objInterCard = New clsInterFaceCard
        Call objInterCard.Init(Nothing, glngSys, glngModul, gcnOracle, gstrDBUser)
    End If
    With rsPatiInfor
        If .State = adStateOpen Then .Close
        '病人ID,姓名,性别,年龄,出生日期,出生地点,身份证号,其他证件,身份,职业,家庭地址,家庭电话,家庭邮编,
        '工作单位,单位邮编,医保号,医疗付款方式,费别,国籍,民族,婚姻状况,区域
        
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, objInterCard.GetPatiInforMaxLen("姓名"), adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 4, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "出生日期", adDate, , adFldIsNullable
        .Fields.Append "出生地点", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "身份证号", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "其他证件", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "身份", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "职业", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "家庭地址", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "家庭电话", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "家庭邮编", adLongVarChar, 6, adFldIsNullable
        .Fields.Append "合同单位ID", adDouble, 18, adFldIsNullable
        .Fields.Append "工作单位", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "单位电话", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "单位邮编", adLongVarChar, 6, adFldIsNullable
        .Fields.Append "医保号", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "医疗付款方式", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "费别", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "国籍", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "民族", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "婚姻状况", adLongVarChar, 4, adFldIsNullable
        .Fields.Append "区域", adLongVarChar, 30, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    zlInitPati = True
End Function

Public Function InitRegist(ByVal lngSys As Long, ByVal lngModul As Long, ByVal cnOracle As ADODB.Connection, ByVal strDbUser As String, _
                Optional objRegist As zlPublicExpense.clsRegist, _
                Optional objExseSvr As zlPublicExpense.clsExpenceSvr, _
                Optional objService As zlPublicExpense.clsService) As Boolean
    '初始化挂号
    Dim strDept As String
    On Error GoTo errH:
    Set objRegist = New clsRegist
    If objRegist.zlInitCommon(lngSys, cnOracle, strDbUser) = False Then Exit Function
    
    Set objExseSvr = New clsExpenceSvr
    If objExseSvr.zlInitCommon(lngSys, lngModul, cnOracle, strDbUser) = False Then Exit Function
    
    Set objService = New zlPublicExpense.clsService
    If objService.zlInitCommon(lngSys, lngModul, cnOracle, strDbUser) = False Then Exit Function
    
    InitRegist = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatiCollect(ByRef cllPati_Out As Collection, ByVal objPati As clsPatientInfo, ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息对像集
    '入参:
    '出参:objPati-病人信息对象集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-29 11:04:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set cllPati_Out = Nothing
    
    If Not objPati Is Nothing Then
        Set cllPati_Out = New Collection
        With objPati
            cllPati_Out.Add .病人ID, "_病人ID"
            cllPati_Out.Add 0, "_主页ID"
            cllPati_Out.Add .姓名, "_姓名"
            cllPati_Out.Add .性别, "_性别"
            cllPati_Out.Add .年龄, "_年龄"
            cllPati_Out.Add .门诊号, "_门诊号"
            cllPati_Out.Add 0, "_住院号"
            cllPati_Out.Add intInsure, "_险类"
        End With
    End If
    If cllPati_Out Is Nothing Then
        Set cllPati_Out = New Collection
        cllPati_Out.Add 0, "_病人ID"
        cllPati_Out.Add 0, "_主页ID"
        cllPati_Out.Add "", "_姓名"
        cllPati_Out.Add "", "_性别"
        cllPati_Out.Add "", "_年龄"
        cllPati_Out.Add "", "_门诊号"
        cllPati_Out.Add "", "_住院号"
        cllPati_Out.Add 0, "_险类"
    End If
    
    GetPatiCollect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
