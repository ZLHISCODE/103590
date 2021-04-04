Attribute VB_Name = "mdlThreeSwap"
Option Explicit

Public Function CheckThreeBalanceToCash(frmMain As Object, ByVal lngModule As Long, _
    cllThreeSwapCards As Collection, ByVal objCard As Card) As Boolean
    '三方卡退现检查
    Dim str操作员 As String
    
    On Error GoTo errHandle
    If Not (objCard.接口序号 > 0 And Not objCard.消费卡) Then CheckThreeBalanceToCash = True: Exit Function
    If CardDelCash(cllThreeSwapCards, objCard.接口序号, objCard) Then CheckThreeBalanceToCash = True: Exit Function
    If CardDefaultCash(cllThreeSwapCards, objCard.接口序号, objCard) = False Then '不允许退现，同时缺省不退现，则不允许强制退现
        ShowMsgbox objCard.名称 & "不允许强制退为其它结算方式！"
        Exit Function
    End If
    
    If zlstr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "三方退款强制退现") Then
        If MsgBox(objCard.名称 & "不支持退现，你确定要将其强制退现吗？", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        str操作员 = zlDatabase.UserIdentifyByUser(frmMain, objCard.名称 & "强制退现，权限验证：", _
            glngSys, lngModule, "三方退款强制退现", , True)
        If str操作员 = "" Then Exit Function
    End If
    CheckThreeBalanceToCash = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CardDelCash(objThreeSwapCards As Collection, _
    ByVal lng卡类别ID As Long, Optional objCard As Card) As Boolean
    '医疗卡是否退现
    'Array(卡类别ID,允许退现,缺省退现,缺省退现方式)
    Dim i As Long
    
    If Not objThreeSwapCards Is Nothing Then
        For i = 1 To objThreeSwapCards.Count
            If objThreeSwapCards(i)(0) = lng卡类别ID Then
                CardDelCash = objThreeSwapCards(i)(1)
                Exit Function
            End If
        Next
    End If
    
    If Not objCard Is Nothing Then
        CardDelCash = objCard.是否退现
    End If
End Function

Public Function CardDefaultCash(objThreeSwapCards As Collection, _
    ByVal lng卡类别ID As Long, Optional objCard As Card) As Boolean
    '医疗卡是否缺省退现
    'Array(卡类别ID,允许退现,缺省退现,缺省退现方式)
    Dim i As Long
    
    If Not objThreeSwapCards Is Nothing Then
        For i = 1 To objThreeSwapCards.Count
            If objThreeSwapCards(i)(0) = lng卡类别ID Then
                CardDefaultCash = objThreeSwapCards(i)(2)
                Exit Function
            End If
        Next
    End If
    
    If Not objCard Is Nothing Then
        CardDefaultCash = objCard.是否缺省退现
    End If
End Function

Public Function CardDefaultBalance(objThreeSwapCards As Collection, ByVal lng卡类别ID As Long) As String
    '医疗卡缺省退现方式
    'Array(卡类别ID,允许退现,缺省退现,缺省退现方式)
    Dim i As Long
    
    If Not objThreeSwapCards Is Nothing Then
        For i = 1 To objThreeSwapCards.Count
            If objThreeSwapCards(i)(0) = lng卡类别ID Then
                CardDefaultBalance = objThreeSwapCards(i)(3)
                Exit Function
            End If
        Next
    End If
End Function

Public Function CheckDelToCash(objThreeSwap As clsThreeSwap, cllThreeSwapCards As Collection, ByVal lng卡类别ID As Long, _
    ByVal str结算方式 As String, ByVal dblMoney As Double, ByVal lng结算序号 As Long, _
    ByVal str卡号 As String, ByVal str交易流水号 As String, ByVal str交易说明 As String) As Boolean
    '三方卡退现检查
    'Array(卡类别ID,允许退现,缺省退现,缺省退现方式)
    Dim strXMLExpend As String, bln允许退现 As Boolean
    Dim bln缺省退现  As Boolean, str缺省退现方式  As String
    
    On Error GoTo ErrHandler
    If cllThreeSwapCards Is Nothing Then Set cllThreeSwapCards = New Collection
    If CollExitsValue(cllThreeSwapCards, "K" & lng卡类别ID) Then CheckDelToCash = True: Exit Function
    
    strXMLExpend = _
        "<INPUT>" & vbCrLf & _
        "  <TKLIST>" & vbCrLf & _
        "    <TK>" & vbCrLf & _
        "      <TKFS>" & str结算方式 & "</TKFS>" & vbCrLf & _
        "      <TKJE>" & dblMoney & "</TKJE>" & vbCrLf & _
        "      <JYLSH>" & str交易流水号 & "</JYLSH>" & vbCrLf & _
        "      <JYSM>" & str交易说明 & "</JYSM>" & vbCrLf & _
        "    </TK>" & vbCrLf & _
        "  </TKLIST>" & vbCrLf & _
        "</INPUT>"
    
    bln允许退现 = objThreeSwap.CheckDelToCash(Val("7-消费卡收款"), _
        lng结算序号, lng卡类别ID, str卡号, str交易流水号, str交易说明, dblMoney, _
        strXMLExpend, bln缺省退现, str缺省退现方式)
    
    cllThreeSwapCards.Add Array(lng卡类别ID, bln允许退现, bln缺省退现, str缺省退现方式), "K" & lng卡类别ID
    
    CheckDelToCash = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function zlGetCorrectCardSql(ByVal strNO As String, ByVal str卡号 As String, ByVal lng病人ID As Long, _
                ByVal lng结帐ID As Long, ByVal str结算方式 As String, ByVal dbl结算金额 As Double, _
                ByVal str发卡时间 As String, Optional ByVal lng三方卡ID As Long, Optional ByVal bln消费卡 As Boolean, _
                Optional ByVal str支付卡号 As String) As String
    Dim strSQL As String
    
    strSQL = "zl_医疗卡记录_结算校正("
    '单据号_In         病人预交记录.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '实际票号_In       住院费用记录.实际票号%Type,
    strSQL = strSQL & "'" & str卡号 & "',"
    '病人id_In         住院费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '结帐id_In         住院费用记录.结帐id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '结算信息_In       Varchar2,
    strSQL = strSQL & "'" & IIf(lng三方卡ID > 0, str结算方式, "") & "',"
    '结算金额_In       病人预交记录.冲预交%Type,
    strSQL = strSQL & "" & dbl结算金额 & ","
    '发卡时间_In       住院费用记录.登记时间%Type,
    strSQL = strSQL & "" & str发卡时间 & ","
    '操作员编号_In     病人预交记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '操作员姓名_In     病人预交记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '卡类别id_In       病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(lng三方卡ID > 0 And Not bln消费卡, lng三方卡ID, "NULL") & ","
    '结算卡序号_In     病人预交记录.结算卡序号%Type := Null,
    strSQL = strSQL & "" & IIf(lng三方卡ID > 0 And bln消费卡, lng三方卡ID, "NULL") & ","
    '卡号_In           病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & str支付卡号 & "',"
    '交易流水号_In     病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '交易说明_In       病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "" & "NULL" & ")"
    
    zlGetCorrectCardSql = strSQL
End Function



Public Function zlGetUpdateSql(ByVal strNO As String, ByVal lng结帐ID As Long, _
    Optional ByVal str结算方式 As String, Optional ByVal dbl结算金额 As Double, _
    Optional ByVal int完成标志 As Integer, Optional ByVal int校对标志 As Integer = 2, _
    Optional ByVal lng卡类别ID As Long, Optional ByVal bln消费卡 As Boolean, _
    Optional ByVal str卡号 As String, Optional ByVal str交易流水号 As String, _
    Optional ByVal str交易说明 As String, Optional ByVal bln普通结算 As Boolean, _
    Optional ByVal str结算号码 As String, Optional ByVal str结算摘要 As String) As String

    Dim strSQL As String

    'Zl_医疗卡结算_Modify
    strSQL = "Zl_医疗卡结算_Modify("
    '      单据号_In     住院费用记录.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '      结帐id_In     住院费用记录.结帐id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '      结算方式_In       病人预交记录.结算方式%Type := NULL,
    strSQL = strSQL & "'" & str结算方式 & "',"
    '      结算金额_In       病人预交记录.冲预交%Type := 0,
    strSQL = strSQL & "" & dbl结算金额 & ","
    '      完成标志_In       Number := 0,
    strSQL = strSQL & "" & int完成标志 & ","
    '      卡类别ID_In       病人预交记录.卡类别ID%Type := Null,
    strSQL = strSQL & "" & ZVal(lng卡类别ID) & ","
    '      消费卡_In         Number := 0,
    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
    '      卡号_In           病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & str卡号 & "',"
    '      交易流水号_In     病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '      交易说明_In       病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & str交易说明 & "',"
    '      普通结算_In Number:=0
    strSQL = strSQL & "" & IIf(bln普通结算, 1, 0) & ","
    '      结算号码_In       病人预交记录.结算号码%Type := Null,
    strSQL = strSQL & "'" & str结算号码 & "',"
    '      摘要_In           病人预交记录.摘要%Type := Null
    strSQL = strSQL & "'" & str结算摘要 & "',"
    '      校对标志_In       病人预交记录.校对标志%Type := 2
    strSQL = strSQL & "" & int校对标志 & ")"
    zlGetUpdateSql = strSQL
End Function
Public Function GetThirdUpdateSQL(ByVal lng预交ID As Long, ByVal strCardNo As String, ByVal str结算方式 As String, ByVal db金额 As Double, ByVal str结算号码 As String, _
                            ByVal strSwapGlideNO As String, ByVal strSwapMemo As String, ByVal str摘要 As String, ByVal intNormal As Integer, cllThird As Collection, Optional ByVal blnRetrun As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新三方交易信息
    '参数: blnRetrun-是否退款，为true只更新结算方式，及是否普通结算
    '编制:
    '日期:2018-09-28
    '说明:
    '问题:132256
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String

    strSQL = "Zl_病人预交记录_Modify("
    '  Id_In         病人预交记录.Id%Type,
    strSQL = strSQL & "" & lng预交ID & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  结算金额_In   病人预交记录.金额%Type,
    strSQL = strSQL & "" & IIf(blnRetrun, "Null", db金额) & ","
    '  结算号码_In   病人预交记录.结算号码%Type,
    strSQL = strSQL & IIf(blnRetrun, "Null", "'" & str结算号码 & "'") & ","
    '  卡号_In       病人预交记录.卡号%Type,
    strSQL = strSQL & IIf(blnRetrun, "Null", "'" & strCardNo & "'") & ","
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & IIf(blnRetrun, "Null", "'" & strSwapGlideNO & "'") & ","
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & IIf(blnRetrun, "Null", "'" & strSwapMemo & "'") & ","
    '  结算摘要_In   病人预交记录.摘要%Type,
    strSQL = strSQL & IIf(blnRetrun, "Null", "'" & str摘要 & "'") & ","
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  普通结算_In Number:=0
    strSQL = strSQL & "" & intNormal & ")"

    zlAddArray cllThird, strSQL
    GetThirdUpdateSQL = True

End Function


Public Function zlAddUpdateSwapSQL(ByVal bln预交 As Boolean, ByVal strIDs As String, _
    ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByRef str卡号 As String, ByVal str交易流水号 As String, ByVal str交易说明 As String, _
    ByRef cllPro As Collection, Optional ByVal int校对标志 As Integer = 0, _
    Optional ByVal int发送标志 As Integer = 0, Optional ByVal bln消费卡管理 As Boolean, _
    Optional ByVal lng关联交易ID As Long, Optional ByVal strExpend As String, _
    Optional dbl金额 As Double, Optional ByVal str单据号 As String, _
    Optional ByVal bln退费 As Boolean, Optional strErrMsg As String, Optional ByVal dbl总金额 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新三方交易流水号和流水说明
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    '出参:cllPro-返回SQL集
    '编制:刘兴洪
    '日期:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strValue As String, strNont As String
    Dim str结算方式 As String, strNO As String
    Dim str结算号码 As String, str结算摘要 As String, str支付卡号 As String
    Dim dbl结算金额 As Double, dblTotalMoney As Double
    Dim i As Long, lngRow As Long, blnNotFisrt As Boolean, bln普通结算 As Boolean
    On Error GoTo errH
    
    If Not bln消费卡 Then
        If zlXML_ExistNode(strExpend, "OUTPUT") Then
            If dbl总金额 = 0 Then dbl总金额 = dbl金额
            If bln退费 Then
                strNont = "TK"
                Call zlXML_GetChildRows("TKLIST", "TK", lngRow)
            Else
                strNont = "JY"
                Call zlXML_GetChildRows("JYLIST", "JY", lngRow)
            End If
            For i = 0 To lngRow - 1
                '||交易方式,交易金额,交易流水号,交易说明,单据号,普通结算||...
                If blnNotFisrt Then
                    strErrMsg = "卡费目前不支持使用多种支付方式结算，请立即与管理联系核查和处理这部分数据。"
                    Exit Function
                End If
                If bln退费 Then
                    Call zlXML_GetChildNodeValue(strNont, "TKFS", i, 0, strValue)
                    str结算方式 = strValue: strValue = ""
                    Call zlXML_GetChildNodeValue(strNont, "TKJE", i, 0, strValue)
                    dbl结算金额 = Val(strValue): strValue = ""
                    dblTotalMoney = dblTotalMoney + dbl结算金额
                Else
                    Call zlXML_GetChildNodeValue(strNont, "JYFS", i, 0, strValue)
                    str结算方式 = strValue: strValue = ""
                    Call zlXML_GetChildNodeValue(strNont, "JYJE", i, 0, strValue)
                    dbl结算金额 = Val(strValue): strValue = ""
                    dblTotalMoney = dblTotalMoney + dbl结算金额
                    Call zlXML_GetChildNodeValue(strNont, "JSHM", i, 0, strValue)
                    str结算号码 = strValue: strValue = ""
                    Call zlXML_GetChildNodeValue(strNont, "JSZY", i, 0, strValue)
                    str结算摘要 = strValue: strValue = ""
                    Call zlXML_GetChildNodeValue(strNont, "KH", i, 0, strValue)
                    str支付卡号 = IIf(strValue <> "", strValue, str卡号)
                End If
                Call zlXML_GetNodeValue("JYLSH", i, strValue)
                str交易流水号 = strValue
                Call zlXML_GetNodeValue("JYSM", i, strValue)
                str交易说明 = strValue
                
                Call zlXML_GetNodeValue("DJH", i, strValue)
                strNO = strValue
                Call zlXML_GetNodeValue("SFPTJS", i, strValue)
                bln普通结算 = Val(strValue) = 1
            
                If bln预交 Then
                    Call GetThirdUpdateSQL(Val(strIDs), str支付卡号, str结算方式, dbl金额, str结算号码, str交易流水号, str交易说明, str结算摘要, IIf(bln普通结算, 1, 0), cllPro)
                Else
                    strSQL = zlGetUpdateSql(str单据号, Val(strIDs), str结算方式, dbl金额, , , lng卡类别ID, , str支付卡号, str交易流水号, str交易说明, bln普通结算, str结算号码, str结算摘要)
                    zlAddArray cllPro, strSQL
                End If
                blnNotFisrt = True
            Next
            str卡号 = str支付卡号
            If RoundEx(dblTotalMoney, 6) <> RoundEx(dbl总金额, 6) And dbl总金额 <> 0 Then
                strErrMsg = "结算金额:" & dbl总金额 & "元，与实际支付的金额:" & dblTotalMoney & "元不一致。" & vbCrLf & _
                            "请立即与管理联系核查和处理这部分数据!"
                Exit Function
            End If
            zlAddUpdateSwapSQL = True
            Exit Function
        End If
    End If
    
    strSQL = "Zl_三方接口更新_Update("
    '  卡类别id_In   病人预交记录.卡类别id%Type,
    strSQL = strSQL & "" & lng卡类别ID & ","
    '  消费卡_In     Number,
    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
    '  卡号_In       病人预交记录.卡号%Type,
    strSQL = strSQL & "'" & str卡号 & "',"
    '  结帐ids_In    Varchar2,
    strSQL = strSQL & "'" & strIDs & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type
    strSQL = strSQL & "'" & str交易说明 & "',"
    '  预交款缴款_In Number := 0,--1-代表预交款缴款;0-代表消费扣款
    strSQL = strSQL & "" & IIf(bln预交, 1, 0) & ","
    '  退费标志_In   Number := 0,--1-进行退费处理;0-支付处理
    strSQL = strSQL & "0,"
    '  校对标志_In   Number := Null,
    strSQL = strSQL & "" & int校对标志 & ","
    '  发送标志_In   Number := 0,
    strSQL = strSQL & "" & int发送标志 & ","
    '  消费卡管理_In Number := 0 --1-消费卡管理调用，此时 消费卡_IN 肯定为0
    strSQL = strSQL & "" & IIf(bln消费卡管理, 1, 0) & ")"
    zlAddArray cllPro, strSQL
    zlAddUpdateSwapSQL = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlAddThreeSwapSQLToCollection(ByVal bln预交款 As Boolean, _
    ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal str卡号 As String, strExpend As String, ByRef cllPro As Collection, _
    Optional ByVal lng预交ID As Long, Optional ByVal int性质 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存三方结算数据
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    ' 出参:cllPro-返回SQL集
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, varData As Variant, varTemp As Variant, i As Long, lngRow As Long
    Dim str交易信息 As String, strTemp As String, strValue As String
     
    Err = 0: On Error GoTo Errhand:
    '先提交,这样避免风险,再更新相关的交易信息
    'strExpend:交易扩展信息,格式:项目名称|项目内容||...
    If zlXML_Init("OUTPUT") Then
        If zlXML_LoadXMLToDOMDocument(strExpend, False, True) Then
            Call zlXML_GetChildRows("Expends", "Expend", lngRow)
            For i = 0 To lngRow - 1
                Call zlXML_GetNodeValue("XMMC", i, strValue)
                strTemp = strTemp & "||" & strValue
                Call zlXML_GetNodeValue("XMNR", i, strValue)
                strTemp = strTemp & "|" & strValue
            Next
            If strTemp <> "" Then strTemp = Mid(strTemp, 3)
        Else
            strTemp = strExpend
        End If
    Else
        strTemp = strExpend
    End If
    
    varData = Split(strTemp, "||")

    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If zlCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                    str交易信息 = Mid(str交易信息, 3)
                    'Zl_三方结算交易_Insert
                    strSQL = "Zl_三方结算交易_Insert("
                    '卡类别id_In 病人预交记录.卡类别id%Type,
                    strSQL = strSQL & "" & lng卡类别ID & ","
                    '消费卡_In   Number,
                    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
                    '卡号_In     病人预交记录.卡号%Type,
                    strSQL = strSQL & "'" & str卡号 & "',"
                    '结帐ids_In  Varchar2,
                    strSQL = strSQL & "'" & strIDs & "',"
                    '交易信息_In Varchar2:交易项目|交易内容||...
                    strSQL = strSQL & "'" & str交易信息 & "',"
                    '预交款缴款_In Number := 0
                    strSQL = strSQL & IIf(bln预交款, "1", "0") & ","
                    '结算方式_In   病人预交记录.结算方式%Type := Null,
                    strSQL = strSQL & "NULL" & ","
                    '预交id_In     病人预交记录.Id%Type := Null,
                    strSQL = strSQL & IIf(lng预交ID = 0, "NULL", lng预交ID) & ","
                    '性质_In       三方结算交易.性质%Type := Null
                    strSQL = strSQL & int性质 & ")"
                    zlAddArray cllPro, strSQL
                    str交易信息 = ""
                End If
                str交易信息 = str交易信息 & "||" & strTemp
            End If
        End If
    Next
    If str交易信息 <> "" Then
        str交易信息 = Mid(str交易信息, 3)
        'Zl_三方结算交易_Insert
        strSQL = "Zl_三方结算交易_Insert("
        '卡类别id_In 病人预交记录.卡类别id%Type,
        strSQL = strSQL & "" & lng卡类别ID & ","
        '消费卡_In   Number,
        strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
        '卡号_In     病人预交记录.卡号%Type,
        strSQL = strSQL & "'" & str卡号 & "',"
        '结帐ids_In  Varchar2,
        strSQL = strSQL & "'" & strIDs & "',"
        '交易信息_In Varchar2:交易项目|交易内容||...
        strSQL = strSQL & "'" & str交易信息 & "',"
        '预交款缴款_In Number := 0
        strSQL = strSQL & IIf(bln预交款, "1", "0") & ","
        '结算方式_In   病人预交记录.结算方式%Type := Null,
        strSQL = strSQL & "NULL" & ","
        '预交id_In     病人预交记录.Id%Type := Null,
        strSQL = strSQL & IIf(lng预交ID = 0, "NULL", lng预交ID) & ","
        '性质_In       三方结算交易.性质%Type := Null
        strSQL = strSQL & int性质 & ")"
        zlAddArray cllPro, strSQL
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

