Attribute VB_Name = "mdlCliniBalance"
Option Explicit

Public Function Get应付款结算方式(Optional ByVal str场合 As String = "收费") As String
    '获取应付款结算方式，退支票
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSql = _
        "Select b.名称" & vbNewLine & _
        "From 结算方式应用 A, 结算方式 B" & vbNewLine & _
        "Where b.名称 = a.结算方式 And a.付款方式 Is Null And Nvl(b.应付款, 0) = 1" & vbNewLine & _
        "      And a.应用场合 = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "mdlCliniBalance", str场合)
    If rsTemp.EOF Then Exit Function
    
    Get应付款结算方式 = Nvl(rsTemp!名称)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub DeleteBalanceRecord(ByVal lng结帐ID As Long, _
    ByVal lng关联交易ID As Long, Optional ByVal lng卡类别ID As Long, _
    Optional ByVal lng结算序号 As Long, Optional ByVal blnMultiDel As Boolean, _
    Optional cllPro As Collection)
    '三方交易调用失败删除结算记录
    '入参：
    '   blnMultiDel 是否多笔退款
    Dim strSql As String
    
    On Error GoTo ErrHandler
    'Zl_病人结算记录_Delete(
    strSql = "Zl_病人结算记录_Delete("
    '  结帐id_In     病人预交记录.结帐id%Type := Null,
    strSql = strSql & "" & ZVal(lng结帐ID) & ","
    '  关联交易id_In 病人预交记录.关联交易id%Type := Null,
    strSql = strSql & "" & ZVal(lng关联交易ID) & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSql = strSql & "" & ZVal(lng卡类别ID) & ","
    '  结算序号_In   病人预交记录.结算序号%Type := Null,
    strSql = strSql & "" & ZVal(lng结算序号) & ","
    '  多笔退款_In   Number := 0
    strSql = strSql & "" & IIf(blnMultiDel, 1, 0) & ")"
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "删除结算记录"
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

Public Sub CancelBillBalance(ByVal lng结帐ID As Long, Optional ByVal strNo As String, _
    Optional cllPro As Collection)
    '取消单据的结算
    Dim strSql As String
    
    On Error GoTo ErrHandler
    'Zl_门诊收费结算_Cancel_S(
    strSql = "Zl_门诊收费结算_Cancel_S("
    '  结帐id_In   门诊费用记录.结帐id%Type,
    strSql = strSql & "" & lng结帐ID & ","
    '  No_In       门诊费用记录.No%Type := Null
    strSql = strSql & "'" & strNo & "')"
    
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

Public Sub CancelBillDelBalance(ByVal lng结帐ID As Long, Optional ByVal lng冲销ID As Long, _
    Optional cllPro As Collection)
    '取消单据的退费
    Dim strSql As String
    
    On Error GoTo ErrHandler
    'Zl_门诊退费结算_Cancel_S(
    strSql = "Zl_门诊退费结算_Cancel_S("
    '  结帐id_In   门诊费用记录.结帐id%Type,
    strSql = strSql & "" & lng结帐ID & ","
    '  重结id_In 门诊费用记录.结帐id%Type := Null
    strSql = strSql & "" & ZVal(lng冲销ID) & ")"
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "取消单据的退费"
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

Public Function UpdateErrBillOperator(ByVal lng病人ID As Long, ByVal lng结算序号 As Long) As Boolean
    '门诊异常收费，更新操作员
    Dim strSql As String
    
    On Error GoTo ErrHandler
    'Zl_门诊异常收费_更新操作员
    strSql = "Zl_门诊异常收费_更新操作员("
    '病人id_In     门诊费用记录.病人id%Type,
    strSql = strSql & "" & lng病人ID & ","
    '操作员编号_In 门诊费用记录.操作员编号%Type,
    strSql = strSql & "'" & UserInfo.编号 & "',"
    '操作员姓名_In 门诊费用记录.操作员姓名%Type,
    strSql = strSql & "'" & UserInfo.姓名 & "',"
    '结算序号_In   病人预交记录.结算序号%Type
    strSql = strSql & lng结算序号 & ")"
    zlDatabase.ExecuteProcedure strSql, "门诊异常收费更新操作员"
    UpdateErrBillOperator = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Public Function Init支付方式(objPayCards As Cards, Optional ByVal blnAdd预交款 As Boolean, _
    Optional ByVal str脱机医保 As String, Optional ByRef str有效脱机医保 As String) As Boolean
    '加载有效的支付方式
    '说明：预交款的结算性质为-99
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim objCard As Card, objCards As Cards, lngKey As Long
    
    str有效脱机医保 = ""
    Set objPayCards = New Cards
    Set objCards = New Cards
    
    Set rsTemp = Get结算方式("收费")
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
        '入参:bytType-  0-所有医疗卡;
        '               1-启用的医疗卡,
        '               2-所有存在三方账户的三方卡
        '               3-启用的三方账户的医疗卡
       Set objCards = gobjSquare.objOneCardComLib.zlGetCards(3)
    End If
    
    With rsTemp
        lngKey = 1
        Do While Not .EOF
            blnFind = False
            For Each objCard In objCards
                If objCard.结算方式 = Nvl(!名称) Then blnFind = True: Exit For
            Next
            
            '性质:1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
            If Not blnFind And InStr("3,4,8", Val(Nvl(!性质))) = 0 And Val(Nvl(!应付款)) = 0 Then
                If InStrEx(str脱机医保, Nvl(!名称), "|") = False Then
                    Set objCard = New Card
                    objCard.短名 = Mid(Nvl(!名称), 1, 1)
                    objCard.接口编码 = Nvl(!编码)
                    objCard.接口程序名 = ""
                    objCard.接口序号 = -1 * lngKey
                    objCard.结算方式 = Nvl(!名称)
                    objCard.名称 = Nvl(!名称)
                    objCard.启用 = True
                    objCard.缺省标志 = Val(Nvl(!缺省)) = 1
                    objCard.支付启用 = True
                    objCard.结算性质 = Val(!性质)
                    objPayCards.Add objCard, "K" & lngKey
                    
                    lngKey = lngKey + 1
                Else
                    str有效脱机医保 = str有效脱机医保 & "|" & Nvl(!名称)
                End If
            End If
            .MoveNext
        Loop
    End With
    If str有效脱机医保 <> "" Then str有效脱机医保 = Mid(str有效脱机医保, 2)
    
    '加三方卡,结算方式要设置了"费用"应用场合才能使用
    For Each objCard In objCards
        rsTemp.Filter = "名称='" & objCard.结算方式 & "'"
        If Not rsTemp.EOF Then
            objCard.缺省标志 = Val(Nvl(rsTemp!缺省)) = 1
            objCard.支付启用 = True
            objPayCards.Add objCard, "K" & lngKey
            lngKey = lngKey + 1
        End If
    Next
    
    If blnAdd预交款 Then
        '加入预交款结算
        Set objCard = New Card
        objCard.短名 = "预"
        objCard.接口编码 = ""
        objCard.接口程序名 = ""
        objCard.接口序号 = -1 * lngKey
        objCard.结算方式 = "预交款"
        objCard.名称 = "预交款"
        objCard.启用 = True
        objCard.缺省标志 = False
        objCard.支付启用 = True
        objCard.结算性质 = "-99"
        objPayCards.Add objCard, "K" & lngKey
    End If
    
    Init支付方式 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteModifyPatiName(ByVal strNos As String, ByVal strName As String) As Boolean
    '修改病人信息
    '入参：
    '   strNos 多个单据用逗号分隔：A001,A002,...
    '   strName 病人新的姓名
    Dim cllPro As New Collection
    Dim strSql As String, arrNo As Variant, i As Long
    
    If strNos = "" Then ExecuteModifyPatiName = True: Exit Function
    arrNo = Split(strNos, ",")
    For i = 0 To UBound(arrNo)
        'Zl_病人费用记录_Update_S(
        strSql = "Zl_病人费用记录_Update_S( "
        '  No_In       门诊费用记录.No%Type,
        strSql = strSql & "'" & arrNo(i) & "',"
        '  记录性质_In 门诊费用记录.记录性质%Type,
        strSql = strSql & "" & 1 & ","
        '  开单人_In   门诊费用记录.开单人%Type,
        strSql = strSql & "" & "NULL" & ","
        '  发生时间_In 门诊费用记录.发生时间%Type,
        strSql = strSql & "" & "NULL" & ","
        '  姓名_In     门诊费用记录.姓名%Type := Null,
        strSql = strSql & "'" & strName & "')"
        '  来源_In     Integer := 1,--门诊;2-住院
        '  年龄_In     门诊费用记录.年龄%Type := Null,
        '  性别_In     门诊费用记录.性别%Type := Null,
        '  出生日期_In 病人信息.出生日期%Type := Null
        zlAddArray cllPro, strSql
    Next

    On Error GoTo ErrHandler:
    zlExecuteProcedureArrAy cllPro, "修改病人信息"
    ExecuteModifyPatiName = True
    Exit Function
ErrHandler:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
