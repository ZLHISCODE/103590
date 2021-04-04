Attribute VB_Name = "mdlThirdSwap"
Option Explicit

Public Function ZlGetForceDelToCashNote(ByRef cllForceDelToCash As Collection) As String
    '获取强制退现摘要，存入"交易说明"自段中，格式：XXXX强制退现:XXX卡;XXX卡
    '入参：
    '   cllForceDelToCash Array(操作员,卡类别名称)
    Dim str操作员 As String
    Dim strTemp As String, i As Integer
    
    On Error GoTo ErrHandler
    If cllForceDelToCash Is Nothing Then Exit Function
    If cllForceDelToCash.Count = 0 Then Exit Function
    
    str操作员 = cllForceDelToCash(1)(0)
    For i = 1 To cllForceDelToCash.Count
        strTemp = strTemp & ";" & cllForceDelToCash(i)(1)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    ZlGetForceDelToCashNote = str操作员 & "强制退现：" & strTemp
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlThreeBalanceCheck(frmMain As Form, ByVal lngModule As Long, _
    ByVal objCard As Card, ByRef cllForceDelToCash As Collection, _
    ByVal str卡类别名称 As String, ByVal bln允许退现 As Boolean, _
    Optional ByRef bln强制退现 As Boolean, _
    Optional ByVal bln缺省退现 As Boolean) As Boolean
    '三方卡强制退现检查
    '入参：
    '   objCard 医疗卡信息
    '   str卡类别名称 卡类别名称
    '出参：
    '   cllForceDelToCash 强制退现信息：Array(操作员,卡类别名称)
    '返回：允许强制退现，返回True；否则，返回False
    '105432
    Dim str操作员 As String
    
    On Error GoTo ErrHandler
    bln强制退现 = False
    If cllForceDelToCash Is Nothing Then Set cllForceDelToCash = New Collection
    
    If objCard Is Nothing Then
        If bln允许退现 = False And bln缺省退现 = False Then
            ShowMsgbox "未找到『" & str卡类别名称 & "』，" & _
                "无法判断其是否支持退现，不能强制退为其它结算方式！"
            Exit Function
        Else
            If MsgBox("未找到『" & str卡类别名称 & "』，无法判断其是否支持退现，" & _
                "你确定要强制退为其它结算方式吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    Else
        If Not (objCard.接口序号 > 0 And Not objCard.消费卡) Then
            ZlThreeBalanceCheck = True: Exit Function
        End If
        If bln允许退现 Then ZlThreeBalanceCheck = True: Exit Function
        If bln缺省退现 Then '不允许退现，同时缺省不退现，则不允许强制退现
            ZlThreeBalanceCheck = True: Exit Function
        Else
            ShowMsgbox "『" & str卡类别名称 & "』不允许强制退为其它结算方式！"
            Exit Function
        End If
    End If
    
    If zlstr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "三方退款强制退现") Then
        If MsgBox("『" & str卡类别名称 & "』不支持退现，你确定要将其强制退现吗？", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        cllForceDelToCash.Add Array(UserInfo.姓名, str卡类别名称)
    Else
        str操作员 = zlDatabase.UserIdentifyByUser(frmMain, _
            "『" & str卡类别名称 & "』强制退现，权限验证：", _
            glngSys, lngModule, "三方退款强制退现", , True)
        If str操作员 = "" Then Exit Function
        cllForceDelToCash.Add Array(str操作员, str卡类别名称)
    End If
    bln强制退现 = True
    ZlThreeBalanceCheck = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelFeeDetailXML(ByVal lng结帐ID As Long) As String
    '获取传入三方卡退费接口zlRetuenCheck中费用列表
    '入参：
    '   lng结帐ID - 结帐ID
    '返回：
    '      <TFLIST> //退费列表
    '        <NO></NO> // 退费单据
    '        <TFITEM> //退费项
    '          <SerialNum></SerialNum> //序号
    '          …
    '        </TFITEM>
    '      </TFLIST>
    '      ...
    Dim strPriorNO As String
    Dim strSQL As String, rsRecord As ADODB.Recordset
    Dim strXML As String, strXMLSub As String
    
    On Error GoTo ErrHandler
    
    strSQL = _
        "Select a.NO, a.序号, a.实收金额" & vbNewLine & _
        "From 门诊费用记录 A" & vbNewLine & _
        "Where a.结帐id = [1]" & vbNewLine & _
        "Order By a.NO, a.序号"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "mdlPublicThreeSwap", lng结帐ID)
    If rsRecord.RecordCount = 0 Then Exit Function
    
    strXML = "": strPriorNO = ""
    Do While Not rsRecord.EOF
        If strPriorNO <> Nvl(rsRecord!NO) Then
            If strPriorNO <> "" Then
                strXML = strXML & "    </TFITEM>" & vbCrLf
                strXML = strXML & "  </TFLIST>" & vbCrLf
            End If
            strXML = strXML & "  <TFLIST>" & vbNewLine '退费列表
            strXML = strXML & "    <NO>" & Nvl(rsRecord!NO) & "</NO>" & vbCrLf '退费单据
            strXML = strXML & "    <TFITEM>" & vbCrLf '退费项
        End If
        
        strXML = strXML & "      <SerialNum>" & Val(Nvl(rsRecord!序号)) & "</SerialNum>" & vbCrLf '序号
        strPriorNO = Nvl(rsRecord!NO)
        
        rsRecord.MoveNext
    Loop
    
    strXML = strXML & "    </TFITEM>" & vbCrLf
    strXML = strXML & "  </TFLIST>" & vbCrLf
    
    ZlGetDelFeeDetailXML = strXML
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlMakeDelFeeRecord() As ADODB.Recordset
    '构建退费明细
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandler
    rsTmp.Fields.Append "NO", adVarChar, 8, adFldIsNullable
    rsTmp.Fields.Append "序号", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set ZlMakeDelFeeRecord = rsTmp
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelFeeRecord(ByVal lng结帐ID As Long) As ADODB.Recordset
    '获取本次退费项目
    Dim i As Integer
    Dim rsDelFeeRecord As ADODB.Recordset
    Dim strSQL As String, rsRecord As ADODB.Recordset
    
    On Error GoTo ErrHandler
    Set rsDelFeeRecord = ZlMakeDelFeeRecord()
    
    strSQL = _
        "Select a.NO, a.序号, a.实收金额" & vbNewLine & _
        "From 门诊费用记录 A" & vbNewLine & _
        "Where a.结帐id = [1]" & vbNewLine & _
        "Order By a.NO, a.序号"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "mdlPublicThreeSwap", lng结帐ID)
    If rsRecord.RecordCount = 0 Then Exit Function
    
    Do While Not rsRecord.EOF
        With rsDelFeeRecord
            .AddNew
            !NO = Nvl(rsRecord!NO)
            !序号 = Nvl(rsRecord!序号)
            !实收金额 = Nvl(rsRecord!实收金额)
            .Update
        End With
        
        rsRecord.MoveNext
    Loop
    
    Set ZlGetDelFeeRecord = rsDelFeeRecord
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetBalanceXML(rsBalance As ADODB.Recordset, _
    rsBalanceByNo As ADODB.Recordset, _
    ByVal lng原结帐ID As Long, ByVal lng卡类别ID As Long, _
    Optional ByRef dblMoneyTotal As Double, _
    Optional ByVal bln全退 As Boolean, _
    Optional ByVal rsDelFeeRecord As ADODB.Recordset, _
    Optional ByRef lng关联交易ID As Long, _
    Optional ByVal dblDelMoney As Double) As String
    '获取原始结算信息
    '入参：
    '   rsBalance - 结算数据，来自于病人预交记录，不含异常状态的数据
    '   rsBalanceByNo - 分单据结算数据，来自于医保结算明细，不含异常状态的数据
    '   lng原结帐ID - 原始结帐ID
    '   lng卡类别ID - 医疗卡类别ID
    '   bln全退 - 是否全退
    '   rsDelFeeRecord - 本次退款信息：NO,序号,实收金额；非原样退时，传入
    '   dblDelMoney - 本次实际录入退款金额；非原样退时，传入
    '出参：
    '   lng关联交易ID - 本次退款原关联交易ID
    '   dblMoneyTotal - 本次退款合计
    '返回：
    '  <TKLIST>//退款列表（35.90以前无此内容）
    '    <TK>
    '      <TKFS>退款方式</TKFS>
    '      <TKJE>退款金额</TKJE>
    '      <JYLSH>原交易流水号</JYLSH>
    '      <JYSM><原交易说明</JYSM>
    '      <DJH>单据号</DJH>
    '    </TK>
    '    …
    '  </TKLIST>
    Dim strXML As String, dblCurMoney As Double
    Dim bln分单据 As Boolean, dblMoney As Double
    Dim i As Integer, j As Integer
    Dim cllNo As Collection, cllBalance As Collection, strKey As String
    Dim rsDelFeeRecordByNo As ADODB.Recordset
    Dim blnFind As Boolean
    
    On Error GoTo ErrHandler
    If rsBalance Is Nothing Then Exit Function
    
    dblMoneyTotal = 0
    If bln全退 Then
        If Not rsDelFeeRecord Is Nothing Then Set rsDelFeeRecordByNo = rsDelFeeRecord
        Set rsDelFeeRecord = ZlMakeDelFeeRecord()
    Else
        If rsDelFeeRecord Is Nothing Then Exit Function
    End If
    
    rsBalance.Filter = "结帐ID=" & lng原结帐ID & " And 卡类别ID=" & lng卡类别ID
    If rsBalance.RecordCount = 0 Then
        '1.可能只是医保进行了重收
        '2.含医保或多种结算方式部分退时都会进行重收
        rsBalance.Filter = "类型=" & Enum_BalanceType.一卡通 & " And 卡类别ID=" & lng卡类别ID & " And 退费=0"
        If rsBalance.RecordCount = 0 Then Exit Function
        lng原结帐ID = Val(Nvl(rsBalance!结帐ID))
    End If
    
    lng关联交易ID = Val(Nvl(rsBalance!关联交易ID))
    
    If Not rsBalanceByNo Is Nothing Then
        rsBalanceByNo.Filter = "关联交易ID=" & lng关联交易ID & " And 卡类别ID=" & lng卡类别ID
        bln分单据 = Not rsBalanceByNo.EOF
    Else
        bln分单据 = False
    End If
    
    dblMoney = 0: strXML = ""
    If bln全退 Then
        '1.全退，不分单据
        If bln分单据 = False Then
            With rsBalance
                .Filter = "关联交易ID=" & lng关联交易ID & " And 卡类别ID=" & lng卡类别ID
                Do While Not .EOF
                    dblMoney = dblMoney + Val(Nvl(!冲预交))
                    
                    rsDelFeeRecord.AddNew
                    rsDelFeeRecord!实收金额 = Val(Nvl(!冲预交))
                    rsDelFeeRecord.Update
                    
                    .MoveNext
                Loop
            End With
        Else '2.全退，分单据(可能是部分单据全退)
            With rsBalanceByNo
                .Filter = "关联交易ID=" & lng关联交易ID & " And 卡类别ID=" & lng卡类别ID
                Do While Not .EOF
                    blnFind = True
                    If Not rsDelFeeRecordByNo Is Nothing Then
                        rsDelFeeRecordByNo.Filter = "NO='" & Nvl(!NO) & "'"
                        blnFind = Not rsDelFeeRecordByNo.EOF
                    End If
                    
                    If blnFind Then
                        dblMoney = dblMoney + Val(Nvl(!金额))
                        
                        rsDelFeeRecord.AddNew
                        rsDelFeeRecord!NO = Nvl(!NO)
                        rsDelFeeRecord!实收金额 = Val(Nvl(!金额))
                        rsDelFeeRecord.Update
                    End If
                    
                    .MoveNext
                Loop
            End With
        End If
        dblDelMoney = dblMoney
    End If
    
    '3.部分退，不分单据
    If bln分单据 = False Then
        With rsDelFeeRecord
            .Filter = "": dblMoney = 0
            Do While Not .EOF
                dblMoney = dblMoney + Val(Nvl(!实收金额))
                .MoveNext
            Loop
        End With
        dblMoneyTotal = dblMoney
        
        dblCurMoney = 0
        Set cllBalance = New Collection
        With rsBalance
            .Filter = "关联交易ID=" & lng关联交易ID & " And 卡类别ID=" & lng卡类别ID
            Do While Not .EOF
                strKey = "_" & Nvl(!结算方式)
                If CollectionExitsValue(cllBalance, strKey) Then
                    dblCurMoney = cllBalance(strKey)(1) + Val(Nvl(!冲预交))
                    cllBalance.Remove strKey
                Else
                    dblCurMoney = Val(Nvl(!冲预交))
                End If
                If dblCurMoney <> 0 Then
                    cllBalance.Add Array(Nvl(!结算方式), dblCurMoney), strKey
                End If
                
                .MoveNext
            Loop
        
            For j = 1 To cllBalance.Count
                If dblMoney > cllBalance(j)(1) Then
                    dblCurMoney = cllBalance(j)(1)
                Else
                    dblCurMoney = dblMoney
                End If
                If dblDelMoney < dblCurMoney Then dblCurMoney = dblDelMoney
                If dblCurMoney <= 0 Then Exit For
            
                .Filter = "结帐ID=" & lng原结帐ID & " And 关联交易ID=" & lng关联交易ID & _
                    " And 结算方式='" & cllBalance(j)(0) & "'" & " And 卡类别ID=" & lng卡类别ID
                If .EOF = False Then
                    strXML = strXML & "    <TK>" & vbCrLf
                    strXML = strXML & "      <TKFS>" & cllBalance(j)(0) & "</TKFS>" & vbCrLf
                    strXML = strXML & "      <TKJE>" & dblCurMoney & "</TKJE>" & vbCrLf
                    strXML = strXML & "      <JYLSH>" & Nvl(!交易流水号) & "</JYLSH>" & vbCrLf
                    strXML = strXML & "      <JYSM>" & Nvl(!交易说明) & "</JYSM>" & vbCrLf
                    strXML = strXML & "      <DJH>" & "" & "</DJH>" & vbCrLf
                    strXML = strXML & "    </TK>" & vbCrLf
                End If
                dblMoney = dblMoney - dblCurMoney
                dblDelMoney = dblDelMoney - dblCurMoney
            Next
        End With
        
        If strXML <> "" Then ZlGetBalanceXML = "  <TKLIST>" & vbCrLf & strXML & "  </TKLIST>"
        Exit Function
    End If
    
    '4.部分退，分单据
    Set cllNo = New Collection
    With rsDelFeeRecord
        .Filter = "": dblMoney = 0
        Do While Not .EOF
            dblMoney = dblMoney + Val(Nvl(!实收金额))
            
            strKey = "_" & Nvl(!NO)
            If CollectionExitsValue(cllNo, strKey) Then
                dblCurMoney = cllNo(strKey)(1) + Val(Nvl(!实收金额))
                cllNo.Remove strKey
            Else
                dblCurMoney = Val(Nvl(!实收金额))
            End If
            cllNo.Add Array(Nvl(!NO), dblCurMoney), strKey
            
            .MoveNext
        Loop
    End With
    dblMoneyTotal = dblMoney
    
    For i = 1 To cllNo.Count
        dblMoney = cllNo(i)(1): dblCurMoney = 0
        Set cllBalance = New Collection
        With rsBalanceByNo
            .Filter = "关联交易ID=" & lng关联交易ID & " And No='" & cllNo(i)(0) & "'" & _
                " And 卡类别ID=" & lng卡类别ID
            Do While Not .EOF
                strKey = "_" & Nvl(!结算方式)
                If CollectionExitsValue(cllBalance, strKey) Then
                    dblCurMoney = cllBalance(strKey)(1) + Val(Nvl(!金额))
                    cllBalance.Remove strKey
                Else
                    dblCurMoney = Val(Nvl(!金额))
                End If
                If dblCurMoney <> 0 Then
                    cllBalance.Add Array(Nvl(!结算方式), dblCurMoney), strKey
                End If
                
                .MoveNext
            Loop
        
            For j = 1 To cllBalance.Count
                If dblMoney > cllBalance(j)(1) Then
                    dblCurMoney = cllBalance(j)(1)
                Else
                    dblCurMoney = dblMoney
                End If
                If dblDelMoney < dblCurMoney Then dblCurMoney = dblDelMoney
                If dblCurMoney <= 0 Then Exit For
                
                .Filter = "结帐ID=" & lng原结帐ID & " And 关联交易ID=" & lng关联交易ID & _
                    " And No='" & cllNo(i)(0) & "' And 结算方式='" & cllBalance(j)(0) & "'" & _
                    " And 卡类别ID=" & lng卡类别ID
                If .EOF = False Then
                    strXML = strXML & "    <TK>" & vbCrLf
                    strXML = strXML & "      <TKFS>" & cllBalance(j)(0) & "</TKFS>" & vbCrLf
                    strXML = strXML & "      <TKJE>" & dblCurMoney & "</TKJE>" & vbCrLf
                    strXML = strXML & "      <JYLSH>" & Nvl(!交易流水号) & "</JYLSH>" & vbCrLf
                    strXML = strXML & "      <JYSM>" & Nvl(!交易说明) & "</JYSM>" & vbCrLf
                    strXML = strXML & "      <DJH>" & cllNo(i)(0) & "</DJH>" & vbCrLf
                    strXML = strXML & "    </TK>" & vbCrLf
                End If
                dblMoney = dblMoney - dblCurMoney
                dblDelMoney = dblDelMoney - dblCurMoney
            Next
        End With
    Next
    
    If strXML <> "" Then ZlGetBalanceXML = "  <TKLIST>" & vbCrLf & strXML & "  </TKLIST>"
    
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelFeeRecordFromGrid(ByVal vsfBill As VSFlexGrid, _
    Optional ByVal bln异常结算 As Boolean) As ADODB.Recordset
    '从界面表格中获取本次退费项目
    Dim i As Integer
    Dim rsDelFeeRecord As ADODB.Recordset
    
    On Error GoTo ErrHandler
    Set rsDelFeeRecord = ZlMakeDelFeeRecord()
    
    With vsfBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 Or bln异常结算 Then
                rsDelFeeRecord.AddNew
                rsDelFeeRecord!NO = .TextMatrix(i, .ColIndex("单据号"))
                rsDelFeeRecord!序号 = .RowData(i)
                rsDelFeeRecord!实收金额 = Val(.TextMatrix(i, .ColIndex("实收金额")))
                rsDelFeeRecord.Update
            End If
        Next
    End With
    
    Set ZlGetDelFeeRecordFromGrid = rsDelFeeRecord
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelFeeDetailXMLFromGrid(ByVal vsfBill As VSFlexGrid, _
    Optional ByVal bln异常结算 As Boolean) As String
    '从界面表格中获取传入三方卡退费接口zlRetuenCheck中费用列表
    '入参：
    '   lng结帐ID - 结帐ID
    '返回：
    '      <TFLIST> //退费列表
    '        <NO></NO> // 退费单据
    '        <TFITEM> //退费项
    '          <SerialNum></SerialNum> //序号
    '          …
    '        </TFITEM>
    '      </TFLIST>
    '      ...
    Dim i As Integer
    Dim strXML As String, blnFindSelectItem As Boolean
    Dim strNo As String, strPriorNO As String
    
    On Error GoTo ErrHandler
    strXML = "": blnFindSelectItem = False
    
    With vsfBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 Or bln异常结算 Then
                blnFindSelectItem = True
                strNo = .TextMatrix(i, .ColIndex("单据号"))
                If strNo <> strPriorNO Then
                    If strPriorNO <> "" Then
                        strXML = strXML & "    </TFITEM>" & vbCrLf
                        strXML = strXML & "  </TFLIST>" & vbCrLf
                    End If
                    strXML = strXML & "  <TFLIST>" & vbNewLine '退费列表
                    strXML = strXML & "    <NO>" & strNo & "</NO>" & vbCrLf '退费单据
                    strXML = strXML & "    <TFITEM>" & vbCrLf '退费项
                End If
                strXML = strXML & "      <SerialNum>" & .RowData(i) & "</SerialNum>" & vbCrLf '序号
                strPriorNO = strNo
            End If
        Next
    End With
    If blnFindSelectItem = False Then Exit Function
    
    strXML = strXML & "    </TFITEM>" & vbCrLf
    strXML = strXML & "  </TFLIST>" & vbCrLf
    
    ZlGetDelFeeDetailXMLFromGrid = strXML
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlCheckThreeSwapValied(frmMain As Object, ByVal lngModule As Long, _
    ByVal lng病人ID As Long, ByVal strPatiName As String, ByVal strSex As String, ByVal strOld As String, _
    ByRef objSquareCard As Object, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal str结算方式 As String, ByRef dblMoney As Double, ByVal strNos As String, _
    Optional ByRef strCardNo As String, Optional ByRef strPassWord As String, _
    Optional ByRef dbl帐户余额 As Double, Optional ByRef str新结算方式 As String, _
    Optional ByVal rsClassMoney As ADODB.Recordset, Optional ByVal str费用来源 As String, _
    Optional ByRef cllSquareBalance As Collection) As Boolean
    '功能:三方支付交易检查
    '入参:
    '   str结算方式-当前结算方式
    '   dblMoney-支付金额
    '   strNos-本次支付所涉及的单据
    '   rsClassMoney-费用类别明细(使用消费卡支付时传入)
    '   str费用来源-当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
    '出参：
    '   strCardNo-刷卡卡号
    '   strPassWord-卡号密码
    '   dbl帐户余额-帐户余额
    '   str新结算方式-刷卡接口返回的结算方式及金额，格式：结算方式|结算金额
    '   cllSquareBalance- 费卡支付明细
    '返回:交易合法返回true,否则返回False
    Dim strXMLExpend As String
    Dim str结算方式_Out As String, dbl结算金额_Out As Double
    Dim strExpand As String

    On Error GoTo ErrHandler
    If objSquareCard Is Nothing Then Exit Function
    
    str新结算方式 = ""
    'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean = False, _
        Optional ByVal bln余额不足禁止 As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal bln转预交 As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXMLIn As String = "", _
        Optional ByVal str费用来源 As String, _
        Optional ByVal lng病人ID As Long, _
        Optional ByRef str结算方式_Out As String = "", _
        Optional ByRef dbl结算金额_Out As Double = 0) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:根据指定支付类别,弹出刷卡窗口
        '入参:rsClassMoney:收费类别,金额
        '        lngCardTypeID-为零时,为老一卡通刷卡
        '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
        '       dblBrushTotaled-消费有效,表示已经刷消费卡总额(主要用于多次刷卡)
        '       str上次限制类别-上次刷消费时的限制类别(同次多次刷消费卡时,需要检查本次刷卡类别与上次类别是否一致,不一致不允许刷卡消费)
        '       varSquareBalance- Collection类型,当前已经刷卡的信息(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文 ))
        '       bln预交-是否转预交
        '       blnAllPay-是否费用全支付，true-费用未支付完不能完成结算，false-可以只支付部分并返回
        '       strXMLExpend-三方卡调用XML入参,目前格式如下:
        '       <IN>
        '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
        '       </IN>
        '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
        '       lng病人ID - 病人ID(使用消费卡支付时传入)
        '出参:str限制类别-限制类别(消费卡返回)
        '        lng消费卡ID-消费卡信息.ID(消费卡返回)
        '       strCardNO-返回刷卡的卡号
        '       strPassWord-返回刷卡所对应的密码
        '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
        '       str结算方式_Out-返回的结算方式
        '       dbl结算金额_Out-返回的结算金额
        '返回:成功,返回true,否则返回False
    strXMLExpend = "<IN><CZLX>0</CZLX></IN>"
    If objSquareCard.zlBrushCard(frmMain, lngModule, rsClassMoney, lng卡类别ID, bln消费卡, _
        strPatiName, strSex, strOld, dblMoney, strCardNo, strPassWord, False, True, False, False, _
        cllSquareBalance, False, False, strXMLExpend, str费用来源, lng病人ID, _
        str结算方式_Out, dbl结算金额_Out) = False Then Exit Function
    
    If str结算方式_Out <> "" Then
        If RoundEx(dblMoney, 6) <> Round(dbl结算金额_Out, 6) Then
            MsgBox str结算方式 & " 实际刷卡支付金额(" & Format(dbl结算金额_Out, "0.00") & ")" & _
                "与应付金额(" & Format(dblMoney, "0.00") & ")不等，支付失败！", vbInformation, gstrSysName
            Exit Function
        End If
        str新结算方式 = str结算方式_Out & "|" & dbl结算金额_Out
    End If
    
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal strCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
        ByVal dblMoney As Double, ByVal strNos As String, _
        Optional ByVal strXMLExpend As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:帐户扣款交易检查
        '入参:frmMain-调用的主窗体
        '       lngModule-调用的模块号
        '       strCardTypeID-卡类别ID
        '       strCardNo-卡号
        '       dblMoney-支付金额(退款时为负数)
        '       strNos-本次支付所涉及的单据
        '       strXMLExpend-(XML串:验证密码:自助机用)
        '出参:
        '   strXMLExpend-(XML串:错误信息)
        '返回:扣款合法,返回true,否则返回Flase
        '编制:刘兴洪
        '日期:2011-05-26 16:42:43
        '说明:
        '   在调用扣款前，由于存在Oracle事务问题， 所以再调用扣款交易前， _
        '   先进行数据的合法性检查,以便控制死锁情况。
    If objSquareCard.zlPaymentCheck(frmMain, lngModule, lng卡类别ID, bln消费卡, _
        strCardNo, dblMoney, strNos) = False Then Exit Function
    
    'zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
        strExpand As String, dblMoney As Double, _
        Optional bln消费卡 As Boolean = False) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:读取帐户余额
        '入参:frmMain-调用的主窗体
        '        lngModule-模块号
        '       strCardNo-卡号
        '       strExpand-预留，为空,以后扩展
        '       bln消费卡-是否为消费卡
        '出参:dblMoney-返回帐户余额
        '返回:函数返回    True:调用成功,False:调用失败
        '编制:刘兴洪
        '日期:2011-05-26 16:29:48
        '说明:
        '       在所有需要扣款的地方，都要检查帐户余额是否充足,帐户不充足时不允许扣款.
        '       如果某些第三方接口不存在余额接口，可以固定返回一定的金额。
    If objSquareCard.zlGetAccountMoney(frmMain, lngModule, _
        lng卡类别ID, strCardNo, strExpand, dbl帐户余额, bln消费卡) = False Then Exit Function
    If dbl帐户余额 <> 0 And dbl帐户余额 < dblMoney Then
        MsgBox str结算方式 & " 帐户余额不足！", vbInformation, gstrSysName
        Exit Function
    End If

    ZlCheckThreeSwapValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLGetThreeSwapXMLExpend(ByVal strXMLExpend As String, ByRef dblOutMoney As Double, _
    ByRef cllBalance As Collection, ByRef strExpend As String) As Boolean
    '功能：解析三方支付数据
    '入参：
    '   strXMLExpend:XML串
    '    <OUTPUT>
    '        <JYLIST> //交易列表
    '            <JY> //保存到预交记录时，按交易流水号及交易说明汇总处理
    '                <JYFS>交易方式</JYFS> //交易方式:即结算方式.名称
    '                <JYJE>交易金额</JYJE>
    '                <JYLSH>交易流水号</JYLSH>
    '                <JYSM>交易说明</JYSM>
    '                <DJH>单据号</DJH> //单据号,多单据收费时有用 ，存储在"医保结算明细"表中,主要是分单据保存
    '                <SFPTJS>是否普通结算</SFPTJS> //是否普通结算(1-普通结算;0-一卡通结算):为1时，在预交记录中不填写卡类别ID,不属于一卡通结算
    '            </JY>
    '            ...
    '        </JYLIST>
    '        <Expends> //交易扩展信息
    '            <Expend> //保存到预交记录时，按交易流水号及交易说明汇总处理
    '                <XMMC>项目名称</XMMC> //交易方式:即结算方式.名称
    '                <XMNR>项目内容</XMNR>
    '            </Expend>
    '            ...
    '        </Expends>
    '    </OUTPUT>
    '出参：
    '   dblOutMoney - 实际支付金额
    '   cllBalance - 结算数据，格式：Array("结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算",交易流水号,交易说明)
    '   strExpend - 扩展数据，格式:项目名称1|项目内容2||…||项目名称n|项目内容n
    Dim lngCount As Long, strValue As String
    Dim i As Integer, strBalance As String
    Dim str交易流水号 As String, str交易说明 As String
    
    On Error GoTo ErrHandler
    dblOutMoney = 0
    Set cllBalance = New Collection: strExpend = ""
    If zlXML_Init() = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
    '结算信息
    Call zlXML_GetRows("JYLIST/JY", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("JYFS", i, strValue)
        strBalance = strValue   '结算方式
        Call zlXML_GetNodeValue("JYJE", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '结算金额
        dblOutMoney = dblOutMoney + Val(strValue)
        strBalance = strBalance & "|" & " " '结算号码
        strBalance = strBalance & "|" & " " '结算摘要
        Call zlXML_GetNodeValue("DJH", i, strValue)
        strBalance = strBalance & "|" & IIf(strValue = "", " ", strValue) '单据号
        Call zlXML_GetNodeValue("SFPTJS", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '是否普通结算
        
        Call zlXML_GetNodeValue("JYLSH", i, strValue)
        str交易流水号 = strValue '交易流水号
        Call zlXML_GetNodeValue("JYSM", i, strValue)
        str交易说明 = strValue   '交易说明
        
        cllBalance.Add Array(strBalance, str交易流水号, str交易说明)
    Next
    
    '扩展信息
    Call zlXML_GetRows("Expends/Expend", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("XMMC", i, strValue)
        strExpend = strExpend & "||" & strValue '项目名称
        Call zlXML_GetNodeValue("XMNR", i, strValue)
        strExpend = strExpend & "|" & strValue '项目内容
    Next
    If strExpend <> "" Then strExpend = Mid(strExpend, 3)
    ZLGetThreeSwapXMLExpend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetThreeSwapBalanceSQL(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    ByVal lng医疗卡类别ID As Long, ByVal byt操作类型 As Byte, _
    ByVal str刷卡卡号 As String, ByVal str结算方式 As String, _
    Optional ByVal lng关联交易ID As Long, Optional ByVal bln删除原结算 As Boolean, _
    Optional ByVal str交易流水号 As String, Optional ByVal str交易说明 As String, _
    Optional ByVal byt校对标志 As Byte = 1) As String
    '获取支付结算SQL
    'byt操作类型 1-三方卡结算,3-消费卡结算,4-三方卡多种结算方式结算
    Dim strSQL As String
    
    ' Zl_门诊收费结算_Modify
    strSQL = "Zl_门诊收费结算_Modify("
    '  --   0-普通收费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退支票额_In:传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②退支票额_In:传入零
    '  --   3-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."  消费卡ID:为零时,根据卡号自动定位
    '  --     ②退支票额_In:传入零
    '  --   4-三方卡结算，多种结算方式:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算"
    '  --     ②退支票额_In:传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    strSQL = strSQL & byt操作类型 & ","
    '    病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & lng病人ID & ","
    '    结帐id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & lng结帐ID & ","
    '    结算方式_In   Varchar2,
    strSQL = strSQL & "'" & str结算方式 & "',"
    '    冲预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '    退支票额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(lng医疗卡类别ID = 0, "NULL", lng医疗卡类别ID) & ","
    '    卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & str刷卡卡号 & "',"
    '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '    交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & str交易说明 & "',"
    '    缴款_In       病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '    找补_In       病人预交记录.找补%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '    误差金额_In   门诊费用记录.实收金额%Type := Null,
    '    -- 误差金额_In:存在误差费时,传入
    strSQL = strSQL & "" & "NULL" & ","
    '  完成结算_In      Number := 0,
    '    -- 完成结算_In:1-完成收费;0-未完成收费
    strSQL = strSQL & "" & 0 & ","
    '  缺省结算方式_In  结算方式.名称%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  冲预交病人ids_In Varchar2 := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  更新交款余额_In  Number := 1,
    strSQL = strSQL & "" & 1 & ","
    '  关联交易id_In    病人预交记录.关联交易id%Type := Null
    strSQL = strSQL & IIf(lng关联交易ID = 0, "NULL", lng关联交易ID) & ","
    '  删除原结算_In    Number := 0,
    strSQL = strSQL & "" & IIf(bln删除原结算, "1", "0") & ","
    '  校对标志_In      病人预交记录.校对标志%Type := 0
    strSQL = strSQL & "" & byt校对标志 & ")"
    ZlGetThreeSwapBalanceSQL = strSQL
End Function

Public Function ZlCheckThreeSwapDelValied(frmMain As Object, ByVal lngModule As Long, _
    ByVal strPatiName As String, ByVal strSex As String, ByVal strOld As String, _
    ByRef objSquareCard As Object, ByVal lng卡类别ID As Long, _
    ByVal blnTransfer As Boolean, ByVal dblMoney As Double, ByVal str原结帐ID As String, _
    Optional ByRef strCardNo As String, Optional ByRef strPassWord As String, _
    Optional ByVal str交易流水号 As String, Optional ByVal str交易说明 As String, _
    Optional ByVal strXMLExpend As String, Optional ByVal bln是否退款验卡 As Boolean) As Boolean
    '功能:三方退款交易检查,不含消费卡
    '入参:
    '     dblMoney-退款金额
    '返回:交易合法返回true,否则返回False
    Dim strBalanceIDs As String
    Dim strExpend As String
    
    On Error GoTo ErrHandler
    If objSquareCard Is Nothing Then Exit Function
    
    '转账模式
    If blnTransfer Then
        'zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln消费卡 As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByRef dbl金额 As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln退费 As Boolean = False, _
            Optional ByRef blnShowPatiInfor As Boolean = False, _
            Optional ByRef bln退现 As Boolean = False, _
            Optional ByVal bln余额不足禁止 As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal bln转预交 As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXMLIn As String = "", _
            Optional ByVal str费用来源 As String, _
            Optional ByVal lng病人ID As Long, _
            Optional ByRef str结算方式_Out As String = "", _
            Optional ByRef dbl结算金额_Out As Double = 0) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:根据指定支付类别,弹出刷卡窗口
            '入参:rsClassMoney:收费类别,金额
            '        lngCardTypeID-为零时,为老一卡通刷卡
            '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
            '       dblBrushTotaled-消费有效,表示已经刷消费卡总额(主要用于多次刷卡)
            '       str上次限制类别-上次刷消费时的限制类别(同次多次刷消费卡时,需要检查本次刷卡类别与上次类别是否一致,不一致不允许刷卡消费)
            '       varSquareBalance- Collection类型,当前已经刷卡的信息(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文 ))
            '       bln预交-是否转预交
            '       blnAllPay-是否费用全支付，true-费用未支付完不能完成结算，false-可以只支付部分并返回
            '       strXMLExpend-三方卡调用XML入参,目前格式如下:
            '       <IN>
            '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
            '       </IN>
            '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
            '       lng病人ID - 病人ID(使用消费卡支付时传入)
            '出参:str限制类别-限制类别(消费卡返回)
            '        lng消费卡ID-消费卡信息.ID(消费卡返回)
            '       strCardNO-返回刷卡的卡号
            '       strPassWord-返回刷卡所对应的密码
            '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
            '       str结算方式_Out-返回的结算方式
            '       dbl结算金额_Out-返回的结算金额
            '返回:成功,返回true,否则返回False
        strExpend = "<IN><CZLX>1</CZLX></IN>"
        If objSquareCard.zlBrushCard(frmMain, lngModule, Nothing, lng卡类别ID, False, _
            strPatiName, strSex, strOld, dblMoney, strCardNo, strPassWord, True, True, False, True, _
            Nothing, False, False, strExpend) = False Then Exit Function

        'zlTransferAccountsCheck(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
            ByVal strCardNo As String, ByVal dblMoney As Double, ByVal strBalanceID As String, _
            Optional ByRef strXMLExpend As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:转帐检查
            '入参:
            '   frmMain-调用的主窗体
            '   lngModule-HIS调用模块号
            '   lngCardTypeID-卡类别ID
            '   strCardNo-卡号
            '   dblMoney-转帐金额(代扣时为负数)
            '   strBalanceID-本次支付结算ID，4-门诊退费业务为原结帐ID
            '   strXMLExpend-XML串:
            '       <IN>
            '           <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务
            '       </IN>
            '出参:strXMLExpend-XML串:
            '        <OUT>
            '           <ERRMSG>错误信息</ERRMSG >
            '        </OUT>
            '返回:检查的数据合法,返回True:否则返回False
            '调用者:医保补充结算(结算时调用)
            '说明:
            '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
            '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
        strExpend = "<IN><CZLX>4</CZLX></IN>"
        If objSquareCard.zltransferAccountsCheck(frmMain, lngModule, lng卡类别ID, _
            strCardNo, dblMoney, str原结帐ID, strExpend) = False Then Exit Function
    Else
        'zlReturncheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
            ByVal strBalanceIDs As String, _
            ByVal dblMoney As Double, ByVal strSwapNo As String, _
            ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:帐户回退交易前的检查
            '入参:frmMain-调用的主窗体
            '       lngModule-调用的模块号
            '       lngCardTypeID-卡类别ID
            '       strCardNo-卡号
            '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算
            '       dblMoney-退款金额
            '       strSwapNo-交易流水号(退款时检查),保险补充结算时传入空
            '       strSwapMemo-交易说明(退款时传入),保险补充结算时传入空
            '       strXMLExpend    XML IN
            '        <TFDATA>   //退费数据
            '            <YCTF>异常退费标志<YCTF> //1-异常重退;0-退费此节点可能没传入
            '            <TFLIST>  //退费列表
            '                <NO></NO>  // 退费单据
            '                <TFITEM>     //退费项
            '                    <SerialNum>序号</SerialNum>
            '                    ….
            '                </ TFITEM >
            '            </TFLIST>
            '
            '            <TKLIST>   //退款列表（35.90以前无此内容）
            '                <TK>
            '                    <TKFS>退款方式</TKFS>// Varchar2    20
            '                    <TKJE>退款金额</TKJE>//NUMBER
            '                    <JSLSH>原交易流水号</JSLSH>//   Varchar2    50
            '                    <JYSM><原交易说明</JYSM>//  Varhcar2    500
            '                    <DJH>单据号</DJH> //    Varchar2    8
            '                </TK>
            '                ....
            '            </TKLIST>
            '        </TFDATA>
            '返回:退款合法,返回true,否则返回Flase
            '说明:
            '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
            '    以便控制死锁情况。
        strBalanceIDs = "3|" & str原结帐ID
        If objSquareCard.zlReturnCheck(frmMain, lngModule, lng卡类别ID, False, strCardNo, _
            strBalanceIDs, dblMoney, str交易流水号, str交易说明, strXMLExpend) = False Then Exit Function
    
        If bln是否退款验卡 Then
           'zlBrushCard(frmMain As Object, _
                ByVal lngModule As Long, _
                ByVal rsClassMoney As ADODB.Recordset, _
                ByVal lngCardTypeID As Long, _
                ByVal bln消费卡 As Boolean, _
                ByVal strPatiName As String, ByVal strSex As String, _
                ByVal strOld As String, ByRef dbl金额 As Double, _
                Optional ByRef strCardNo As String, _
                Optional ByRef strPassWord As String, _
                Optional ByRef bln退费 As Boolean = False, _
                Optional ByRef blnShowPatiInfor As Boolean = False, _
                Optional ByRef bln退现 As Boolean = False, _
                Optional ByVal bln余额不足禁止 As Boolean = True, _
                Optional ByRef varSquareBalance As Variant, _
                Optional ByVal bln转预交 As Boolean = False, _
                Optional ByVal blnAllPay As Boolean = False, _
                Optional ByVal strXMLIn As String = "", _
                Optional ByVal str费用来源 As String, _
                Optional ByVal lng病人ID As Long, _
                Optional ByRef str结算方式_Out As String = "", _
                Optional ByRef dbl结算金额_Out As Double = 0) As Boolean
                '---------------------------------------------------------------------------------------------------------------------------------------------
                '功能:根据指定支付类别,弹出刷卡窗口
                '入参:rsClassMoney:收费类别,金额
                '        lngCardTypeID-为零时,为老一卡通刷卡
                '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
                '       dblBrushTotaled-消费有效,表示已经刷消费卡总额(主要用于多次刷卡)
                '       str上次限制类别-上次刷消费时的限制类别(同次多次刷消费卡时,需要检查本次刷卡类别与上次类别是否一致,不一致不允许刷卡消费)
                '       varSquareBalance- Collection类型,当前已经刷卡的信息(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文 ))
                '       bln预交-是否转预交
                '       blnAllPay-是否费用全支付，true-费用未支付完不能完成结算，false-可以只支付部分并返回
                '       strXMLExpend-三方卡调用XML入参,目前格式如下:
                '       <IN>
                '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
                '       </IN>
                '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
                '       lng病人ID - 病人ID(使用消费卡支付时传入)
                '出参:str限制类别-限制类别(消费卡返回)
                '        lng消费卡ID-消费卡信息.ID(消费卡返回)
                '       strCardNO-返回刷卡的卡号
                '       strPassWord-返回刷卡所对应的密码
                '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
                '       str结算方式_Out-返回的结算方式
                '       dbl结算金额_Out-返回的结算金额
                '返回:成功,返回true,否则返回False
            strExpend = "<IN><CZLX>2</CZLX></IN>"
            If objSquareCard.zlBrushCard(frmMain, lngModule, Nothing, lng卡类别ID, False, _
                strPatiName, strSex, strOld, dblMoney, strCardNo, strPassWord, True, True, False, True, _
                Nothing, False, False, strExpend) = False Then Exit Function
        End If
    End If
    
    ZlCheckThreeSwapDelValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLGetThreeSwapDelXMLExpend(ByVal strXMLExpend As String, ByRef dblOutMoney As Double, _
    ByRef cllBalance As Collection) As Boolean
    '功能：解析三方退款数据
    '入参：
    '   strXMLExpend:XML串
    '    <OUTPUT>
    '        <TKLIST>
    '            <TK>
    '                <TKFS>退款方式</TKFS>
    '                <TKJE>结算金额</TKJE>
    '                <JYLSH>退款交易流水号</JYLSH>
    '                <JYSM>退款交易说明</JYSM>
    '                <DJH>单据号</DJH>
    '                <SFPTJS>是否普通结算</SFPTJS>
    '            </TK>
    '            …
    '        </TKLIST>
    '    </OUTPUT>
    '   blnDelMoney - 是否对金额取相反数
    '出参：
    '   dblOutMoney - 实际退款金额
    '   cllBalance - 结算数据，格式：Array("结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算",交易流水号,交易说明)
    Dim lngCount As Long, strValue As String
    Dim i As Integer, strBalance As String
    Dim str交易流水号 As String, str交易说明 As String
    
    On Error GoTo ErrHandler
    dblOutMoney = 0
    Set cllBalance = New Collection
    If zlXML_Init() = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
    '结算信息
    Call zlXML_GetRows("TKLIST/TK", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("TKFS", i, strValue)
        strBalance = strValue '退款方式
        Call zlXML_GetNodeValue("TKJE", i, strValue)
        strBalance = strBalance & "|" & -1 * Val(strValue)    '结算金额
        dblOutMoney = dblOutMoney + -1 * Val(strValue)
        strBalance = strBalance & "|" & " " '结算号码
        strBalance = strBalance & "|" & " " '结算摘要
        Call zlXML_GetNodeValue("DJH", i, strValue)
        strBalance = strBalance & "|" & IIf(strValue = "", " ", strValue) '单据号
        Call zlXML_GetNodeValue("SFPTJS", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '是否普通结算
        
        Call zlXML_GetNodeValue("JYLSH", i, strValue)
        str交易流水号 = strValue '交易流水号
        Call zlXML_GetNodeValue("JYSM", i, strValue)
        str交易说明 = strValue   '交易说明
        
        cllBalance.Add Array(strBalance, str交易流水号, str交易说明)
    Next
    ZLGetThreeSwapDelXMLExpend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLGetThreeSwapDelBalanceSQL(ByVal lng病人ID As Long, ByVal lng冲销ID As Long, _
    ByVal lng医疗卡类别ID As Long, ByVal byt操作类型 As Byte, _
    ByVal str刷卡卡号 As String, ByVal str结算方式 As String, _
    Optional ByVal lng关联交易ID As Long, Optional ByVal bln删除原结算 As Boolean, _
    Optional ByVal str交易流水号 As String, Optional ByVal str交易说明 As String, _
    Optional ByVal byt校对标志 As Byte = 1) As String
    '获取退款结算SQL
    'byt操作类型 2-三方卡结算,4-消费卡结算,5-三方卡多种结算方式结算
    Dim strSQL As String
    
    'Zl_门诊退费结算_Modify(
    strSQL = "Zl_门诊退费结算_Modify("
    '  --操作类型_In:
    '  --   0-原样退
    '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
    '  --   1-普通退费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --   2.三方卡退费结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --   4-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --   5.三方卡退费结算，多种结算方式:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算"
    '  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    strSQL = strSQL & "" & byt操作类型 & ","
    '  病人id_In        门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  冲销id_In        病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & lng冲销ID & ","
    '  结算方式_In      Varchar2,
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  冲预交_In        病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  卡类别id_In      病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(lng医疗卡类别ID = 0, "NULL", lng医疗卡类别ID) & ","
    '  卡号_In          病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & str刷卡卡号 & "',"
    '  交易流水号_In    病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '  交易说明_In      病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & str交易说明 & "',"
    '  缴款_In          病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  找补_In          病人预交记录.找补%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  误差金额_In      门诊费用记录.实收金额%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  完成退费_In      Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '  原结帐id_In      病人预交记录.结帐id%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  剩余转预交_In    Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '  缺省结算方式_In  结算方式.名称%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  冲预交病人ids_In Varchar2 := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  关联交易id_In    病人预交记录.关联交易id%Type := Null,
    strSQL = strSQL & "" & IIf(lng关联交易ID = 0, "NULL", lng关联交易ID) & ","
    '  删除原结算_In    Number := 0,
    strSQL = strSQL & "" & IIf(bln删除原结算, "1", "0") & ","
    '  校对标志_In      病人预交记录.校对标志%Type := 0
    strSQL = strSQL & "" & byt校对标志 & ")"
    ZLGetThreeSwapDelBalanceSQL = strSQL
End Function

Public Function ZLCheckThreeSwapDelToCash(frmMain As Object, ByVal lngModule As Long, _
    ByRef objSquareCard As Object, rsBalance As ADODB.Recordset, _
    ByVal lng原结帐ID As Long, ByVal lngCardTypeID As Long, _
    ByVal dblMoney As Double, ByVal strXMLExpend As String, _
    Optional blnDelDefaultCash_Out As Boolean, _
    Optional strDefaultDelBalance_Out As String) As Boolean
    '三方结算交易退现检查
    Dim strCardNo As String, lng结帐ID As Long
    Dim strSwapNO As String, strSwapMemo As String
    
    On Error GoTo ErrHandler
    If rsBalance Is Nothing Then Exit Function
    
    rsBalance.Filter = "结帐ID=" & lng原结帐ID & " And 卡类别id=" & lngCardTypeID & " And 退费=0"
    If rsBalance.EOF Then Exit Function
    strCardNo = Nvl(rsBalance!卡号)
    lng结帐ID = Nvl(rsBalance!结帐ID)
    strSwapNO = Nvl(rsBalance!交易流水号)
    strSwapMemo = Nvl(rsBalance!交易说明)
    
    strXMLExpend = "<INPUT>" & vbCrLf & _
                        strXMLExpend & vbCrLf & _
                    "</INPUT>"
    'zlReturnCashCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strXMLExpend As String, _
        Optional blnDelDefaultCash_Out As Boolean, Optional strDefaultDelBalance_Out As String) As Boolean
    '功能:退现交易检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(退款时检查)多种结算方式时，本参数为第一个结算方式的交易流水号
    '       strSwapMemo-交易说明(退款时传入) 多种结算方式时，本参数为第一个结算方式的交易说明
    '       strXMLExpend    XML IN  10.35.90后才支持
    '        <INPUT>
    '            <TKLIST>    //本次退款列表
    '                <TK>
    '                    <TKFS>退款方式</TKFS>
    '                    <TKJE>退款金额</TKJE>
    '                    <JYLSH>原交易流水号</JYLSH>
    '                    <JYSM>原交易说明</JYSM>
    '                </TK>
    '                ....
    '            </TKLIST>
    '        </INPUT>
    '出参:
    '       blnDelDefaultCash_Out-是否缺省退现：接口返回true时有效，true时：表示缺省退成现金（缺省方式为:str缺省退现方式_Out返回值),否则缺省退回原卡，但允许操作员选择退为现金
    '       strDefaultDelBalance_Out-缺省退现方式,比如：支票，现金等
    '       strXMLExpend:10.35.90后才支持
    '        <OUTPUT>
    '            <SFQSTX>是否缺省退现<SFQSTX>//NUMBER 1 是否缺省退现: 1-缺省;0-不缺省，缺省退回原卡，但以许操作员操作退现
    '            <QSTKFS>缺省退现退款方式</QSTKFS>//Varchar2 20 缺省退现退款方式即结算方式.名称
    '                    1.不允许返回三卡方的结算方式
    '                    2.应避免使用：医保类结算，一卡通本身的结算方式和消费卡的一些特殊结算方式。返回这类方式，将被禁使用这些方式
    '        </OUTPUT>
    '返回:退现合法,返回true,否则返回Flase
    If objSquareCard.zlReturnCashCheck(frmMain, lngModule, lngCardTypeID, strCardNo, lng结帐ID, dblMoney, _
        strSwapNO, strSwapMemo, strXMLExpend, _
        blnDelDefaultCash_Out, strDefaultDelBalance_Out) = False Then Exit Function
    
    ZLCheckThreeSwapDelToCash = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
