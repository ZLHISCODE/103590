Attribute VB_Name = "mdl南京市"
Option Explicit
Private mstrPatID As String
Private mobjSystem As New FileSystemObject
Private mobjStream As TextStream
Private mcur个帐余额 As Currency
Public gstr正确姓名 As String
Public gstr人员身份 As String
'Private mdomInput As New MSXML2.DOMDocument
'Private mdomOutput As New MSXML2.DOMDocument
Private mblnInit As Boolean
Public gblnBill As Boolean          '医保票据控制
Public gblnCancel_南京 As Boolean     '如果门诊预算医保返回数据则不允许取消
Public gintBills As Integer
Public glng领用ID As Long
Public glng公用ID As Long
Public gint明细数 As Integer
Public gint收据费目 As Integer
Public gcnNJSYB As New ADODB.Connection

Private Type patInfo_南京市
    医保号 As String
    就诊时间 As String
    病人姓名 As String
    医生编码 As String
    医生姓名 As String
    病种编码 As String
    病种名称 As String
    医保就诊科室码 As String
    医保就诊科室名 As String
    操作人编码 As String
End Type
Public gPatInfo_南京市 As patInfo_南京市

Private Type detailFee_南京市
    行号 As Double
    医保号 As String
    住院序号 As String
    病人姓名 As String
    标志 As String
    费用发生时间 As String
    医院编码 As String
    医院自编码  As String
    医保编码 As String
    名称 As String
    剂量单位 As String
    单价 As Double
    数量 As Double
    操作人编码 As String
    操作人姓名 As String
    产地 As String
    产地特征 As String
    规格 As String
End Type
Private mDetailFee_南京市 As detailFee_南京市

Private Type feeBalance_南京市
    住院序号 As String
    医保卡号 As String
    费用发生时间 As String
    门诊费用合计 As Double
    药费合计 As Double
    治疗项目合计 As Double
    自理费用 As Double
    医保范围费用 As Double
    个人帐户支付 As Double
    统筹支付 As Double
    大病支付 As Double
    个人自付 As Double
    期初个人帐户 As Double
    期末个人帐户 As Double
    操作员编码 As String
    单据号 As String
    险种 As String
    优惠1 As Double
    优惠2 As Double
    优惠3 As Double
    低保帐户支付 As Double
End Type
Public mFeeBalance As feeBalance_南京市

Public gstr备注 As String

Public Function 医保初始化_南京市() As Boolean
    Dim strServer As String, strUser As String, strPass As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    gblnBill = (GetSetting("ZLSOFT", "公共模块\医保票据管理", "医保票据管理", 0) = 1)
    
    If Not mblnInit And gblnBill Then
        strServer = AnalyServer(gcnOracle.ConnectionString)
        Call AnalyConf(strUser, strPass, strServer)
        With gcnNJSYB
            If .State = 1 Then .Close
            .Provider = "MsDataShape"
            .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUser, strPass
            If Err <> 0 Then
                MsgBox "连接中间用户失败！", vbInformation, gstrSysName
                Exit Function
            End If
        End With
        
        '读取明细数与收据费目条数
        gstrSQL = " Select 明细数,收据费目 From 票据打印参数"
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open gstrSQL, gcnNJSYB
        If rsTemp.RecordCount <> 0 Then
            gint明细数 = Nvl(rsTemp!明细数, 0)
            gint收据费目 = Nvl(rsTemp!收据费目, 0)
        End If
        
        If gint明细数 = 0 Or gint收据费目 = 0 Then
            MsgBox "请使用医保票据管理工具设置票据打印参数！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    mblnInit = True
     医保初始化_南京市 = True
End Function

Public Function 身份标识_南京市(Optional bytType As Byte, Optional lng病人ID As Long) As String
    
    On Error GoTo errorhandle
    If bytType = 0 Or bytType = 3 Then
        If gblnBill Then
            '检查是否有自用或公用票据,没有则不允许进行身份验证;因住院在结算时才使用票据,所以不检测
            glng公用ID = GetSetting("ZLSOFT", "公共模块\医保票据管理\门诊", "共用收费票据批次", 0)
            glng领用ID = GetInvoiceGroupID(1, 1, glng领用ID, glng公用ID)
            If glng领用ID <= 0 Then
                Select Case glng领用ID
                    Case 0 '操作失败
                    Case -1
                        MsgBox "你没有自用和共用的医保收费票据,请先领用一批医保票据或设置本地共用医保票据！", vbInformation, gstrSysName
                    Case -2
                        MsgBox "本地共用的医保票据已经用完,请先领用一批医保票据或重新设置本地共用医保票据！", vbInformation, gstrSysName
                End Select
                Exit Function
            End If
        End If
        身份标识_南京市 = frmIdentify南京市.Identify(bytType, lng病人ID)
    Else
        身份标识_南京市 = frm数据交换.getFeeBalance(bytType, lng病人ID)
        Unload frm数据交换
    End If
    
    gblnCancel_南京 = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_南京市(rs明细 As ADODB.Recordset, str结算方式 As String, Optional strAdvance As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '字段：开单人,病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保

    Dim rsTemp As New ADODB.Recordset, curCount As Currency, dbl现金 As Double
    Dim strFile As String, strWrite As String
    Dim strTemp As String
    Dim intOrder As Integer
    Dim dbl实收金额 As Double, str备注 As String
    Dim intSubInsure As Integer, strYHLB As String, dblSubBalance As Double, strSubInsureNO As String, intSubDisable As Integer  '子医保序号，子医保帐户余额及子医保号，停用标志
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
'    Dim InvokeServer As String
    '删除可能存在的前次结算信息文件
    On Error Resume Next
    Call Kill("C:\NJYB\mzjshz.xml")
    
    On Error GoTo errorhandle
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用记录，不能进行结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    '提取当前病人子医保相关信息(保险序号|优惠类别|医保号|余额|停用)
    gstrSQL = " Select 退休证号||'||||' AS 退休证号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前病人子医保相关信息", TYPE_南京市, CLng(rs明细!病人ID))
    intSubInsure = Val(Split(rsTemp!退休证号, "|")(0))
    strYHLB = Split(rsTemp!退休证号, "|")(1)
    strSubInsureNO = Split(rsTemp!退休证号, "|")(2)
    dblSubBalance = Val(Split(rsTemp!退休证号, "|")(3))
    intSubDisable = Val(Split(rsTemp!退休证号, "|")(4)) '停用则不允许使用统筹支付，意思是明细无打折
    Select Case strYHLB
    Case "惠民"
        strYHLB = "1"
    Case "慈善"
        strYHLB = "2"
    Case "零差率"
        strYHLB = "3"
    Case Else
        strYHLB = "0"
    End Select
    
    curCount = 0
    While Not rs明细.EOF
         curCount = curCount + rs明细!实收金额
        rs明细.MoveNext
    Wend
    rs明细.MoveFirst
    
    strAdvance = ""
    If gblnBill Then
        '检查票据
        gintBills = AnalyBill(rs明细)
        If IsEnough() = False Then
            MsgBox "本次收费将使用" & gintBills & "张医保票据，而当前剩余张数不足，请更换票据后重新收费！", vbInformation, gstrSysName
            Exit Function
        End If
        strAdvance = GetNextBill(glng领用ID) '向前台程序返回本次结算所使用的票据开始号码
    End If
    
    '取出病人信息所需内容
    mstrPatID = rs明细!病人ID
    With gPatInfo_南京市
        .就诊时间 = Format(zlDatabase.Currentdate, "yyyyMMddHHmmss")        '得到就诊时间
        .医生姓名 = Nvl(rs明细!开单人)                                               '得到医生姓名
    End With
    
    If Trim(gPatInfo_南京市.医生姓名) = "" Then
        MsgBox "医保病人收费必须输入医生", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "select A.电子邮件 as 医生编码,decode(c.位置,null,c.简码,c.位置) as 医生科室编码,C.名称 as 医生科室名称 from 人员表 A,部门人员 B,部门表 C,临床部门 D " & _
              "where A.id=B.人员id and B.部门id = C.id and C.id=D.部门id and B.缺省=1 and  A.姓名=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医生编码", CStr(rs明细!开单人))
    If rsTemp.EOF Then
        MsgBox "未对应科室的诊疗科目编码,请先正确对应", vbInformation, gstrSysName
        Exit Function
    End If
    
    With gPatInfo_南京市
        .医生编码 = rsTemp!医生编码                                               '取得医生编码
        .医保就诊科室码 = rsTemp!医生科室编码
        .医保就诊科室名 = rsTemp!医生科室名称
        .操作人编码 = UserInfo.编号
    End With
    
    If InitXML = False Then Exit Function
    Set nodRow = InsertChild(mdomInput.documentElement, "RECORD", "")
    Call InsertChild(nodRow, "TBR", gPatInfo_南京市.医保号)
    Call InsertChild(nodRow, "XM", gPatInfo_南京市.病人姓名)
    Call InsertChild(nodRow, "YSM", gPatInfo_南京市.医生编码)
    Call InsertChild(nodRow, "YSXM", gPatInfo_南京市.医生姓名)
    Call InsertChild(nodRow, "BZBM", gPatInfo_南京市.病种编码)
    Call InsertChild(nodRow, "KSM", gPatInfo_南京市.医保就诊科室码)
    Call InsertChild(nodRow, "KSMC", gPatInfo_南京市.医保就诊科室名)
        
    mdomInput.Save "C:\NJYB\mzjzxx.xml"
    
    '取出明细费用所需内容
    gstrSQL = "select 医院编码 from 保险类别 where 序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医院编码", TYPE_南京市)
    If rsTemp.EOF Then
        MsgBox "医院编码未设置,请先设置医院编码", vbInformation, gstrSysName
        Exit Function
    End If
    With mDetailFee_南京市
        .病人姓名 = gPatInfo_南京市.病人姓名
        .费用发生时间 = gPatInfo_南京市.就诊时间
        .医院编码 = rsTemp!医院编码
        .操作人编码 = gPatInfo_南京市.操作人编码
    End With
    
    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称 from (select * from 保险支付项目 where 险类=[1]) A, 收费细目 B where A.收费细目id(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保项目", TYPE_南京市, CLng(rs明细!收费细目ID))
        If IsNull((rsTemp!项目编码)) Then
            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
            Exit Function
        End If
        rs明细.MoveNext
    Loop
    rs明细.MoveFirst
    
    rs明细.MoveFirst
    If InitXML = False Then Exit Function
    Do Until rs明细.EOF
        If rs明细!实收金额 <> 0 Then
            '调用下属子医保得到单条明细的统筹金额
            dbl实收金额 = rs明细!实收金额
            If intSubInsure <> 0 And intSubDisable = 0 Then '低保且帐户未停用
                If Not CreateObject_Insure(intSubInsure, intOrder) Then Exit Function
                If Not gobjInsure_Obj(intOrder).CalcSingleRecord(rs明细!收费细目ID, dbl实收金额, gstr备注, intSubInsure, rs明细.AbsolutePosition) Then Exit Function
            End If
            
            '准备上传数据
            gstrSQL = "select decode(A.类别,'5',0,'6',0,'7',0,1) 标志,A.名称,nvl(b.门诊包装,1) 门诊包装,C.项目编码,a.计算单位,B.产地,decode(B.药品来源,'国产',1,'合资',2,'进口',3,null) 产地特征,B.规格" & _
                  " from 收费细目 A,药品目录 B,保险支付项目 C where A.id = C.收费细目id and A.id=B.药品id(+) and A.id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "收细明细", CLng(rs明细!收费细目ID))
            With mDetailFee_南京市
                .标志 = rsTemp!标志
                .名称 = rsTemp!名称
                .医保编码 = rsTemp!项目编码
                .剂量单位 = Nvl(rsTemp!计算单位)
                .剂量单位 = ToVarchar(.剂量单位, 10)
                .单价 = Val(Format(rs明细!实收金额 / (rs明细!数量 / rsTemp!门诊包装), "#0.0000;-#0.0000;0;"))
                .数量 = rs明细!数量 / rsTemp!门诊包装
                .产地 = Nvl(rsTemp!产地)
                .产地特征 = Nvl(rsTemp!产地特征)
                .规格 = Nvl(rsTemp!规格)
            End With
            
            Set nodRow = InsertChild(mdomInput.documentElement, "RECORD", "")
            Call InsertChild(nodRow, "TBR", gPatInfo_南京市.医保号)
            Call InsertChild(nodRow, "XM", gPatInfo_南京市.病人姓名)
            Call InsertChild(nodRow, "BZ", mDetailFee_南京市.标志)
            Call InsertChild(nodRow, "ZBM", mDetailFee_南京市.医保编码)
            Call InsertChild(nodRow, "SL", mDetailFee_南京市.数量)
            Call InsertChild(nodRow, "DJ", mDetailFee_南京市.单价)
            Call InsertChild(nodRow, "YHLB", strYHLB)
            Call InsertChild(nodRow, "YHJ", Val(Format(dbl实收金额 / (rs明细!数量 / rsTemp!门诊包装), "#0.0000;-#0.0000;0;")))
        End If
          
        rs明细.MoveNext
     Loop

    mdomInput.Save "C:\NJYB\mzcfsj.xml"
    Call DebugTool("明细文件已产生")
    
    '读出医保结算结果
    strTemp = frm数据交换.getFeeBalance
    On Error Resume Next
    Unload frm数据交换
    On Error GoTo errorhandle
    If strTemp = "" Then
        MsgBox "读取医保结算文件过程被中止,无法完成预结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    '取出信息为门诊结算做准备
    '参数的传入
    If InitXML = False Then Exit Function
    Set mdomInput = New MSXML2.DOMDocument
    If mdomInput.Load("c:\njyb\mzjshz.xml") = False Then
        MsgBox "医保服务器返回值格式不正确。", vbInformation, gstrSysName
    Else
        Set nodRowset = mdomInput.documentElement.selectSingleNode("RECORD")
        With mFeeBalance
            .医保卡号 = nodRowset.selectSingleNode("TBR").Text
            .门诊费用合计 = nodRowset.selectSingleNode("ZFY").Text
            .自理费用 = nodRowset.selectSingleNode("GRZL").Text
            .个人帐户支付 = nodRowset.selectSingleNode("ZHZF").Text
            .统筹支付 = nodRowset.selectSingleNode("YBZF").Text
            .个人自付 = nodRowset.selectSingleNode("GRZF").Text
            .单据号 = nodRowset.selectSingleNode("DJH").Text
            If nodRowset.selectSingleNode("FYLB").Text = "门精" Then
                .大病支付 = nodRowset.selectSingleNode("ZFY").Text
            Else
                .大病支付 = 0
            End If
            .险种 = nodRowset.selectSingleNode("XZMC").Text
            .优惠1 = Val(nodRowset.selectSingleNode("YH1").Text)
            .优惠2 = Val(nodRowset.selectSingleNode("YH2").Text)
            .优惠3 = Val(nodRowset.selectSingleNode("YH3").Text)
        End With
    End If
    Call DebugTool("完成各项数据的读取")
    If curCount <> CCur(mFeeBalance.门诊费用合计) Then
        MsgBox "请注意：医保返回费用合计与医院结算费用合计不等" & vbCrLf & _
            "医院：" & curCount & Space(10) & "医保：" & mFeeBalance.门诊费用合计
    End If
    mcur个帐余额 = nodRowset.selectSingleNode("ZHYE").Text
    gstrSQL = "zl_保险帐户_更新信息(" & mstrPatID & "," & TYPE_南京市 & ",'帐户余额','" & mcur个帐余额 & "')"
    Call DebugTool("更新帐户余额:" & gstrSQL)
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

'芮奇
    gstr人员身份 = nodRowset.selectSingleNode("XZMC").Text
    gstrSQL = "zl_保险帐户_更新信息(" & mstrPatID & "," & TYPE_南京市 & ",'人员身份',''" & gstr人员身份 & "'')"
    Call DebugTool("更新人员身份:" & gstrSQL)
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    Call DebugTool("返回结算方式")
    str结算方式 = "个人帐户;" & mFeeBalance.个人帐户支付 & ";0"
    If mFeeBalance.统筹支付 <> 0 Then
        str结算方式 = str结算方式 & "|统筹基金;" & mFeeBalance.统筹支付 & ";0"
    End If
    If mFeeBalance.大病支付 <> 0 Then
        str结算方式 = str结算方式 & "|大病统筹;" & mFeeBalance.大病支付 & ";0"
    End If
    If mFeeBalance.优惠1 <> 0 Then
        str结算方式 = str结算方式 & "|惠民补助;" & mFeeBalance.优惠1 & ";0"
    End If
    If mFeeBalance.优惠2 <> 0 Then
        str结算方式 = str结算方式 & "|慈善减免;" & mFeeBalance.优惠2 & ";0"
    End If
    If mFeeBalance.优惠3 <> 0 Then
        str结算方式 = str结算方式 & "|零差率优惠;" & mFeeBalance.优惠3 & ";0"
    End If
    '计算低保的帐户支付金额
    dbl现金 = curCount - mFeeBalance.个人帐户支付 - mFeeBalance.统筹支付 - mFeeBalance.大病支付 - mFeeBalance.优惠1 - mFeeBalance.优惠2 - mFeeBalance.优惠3
    If dbl现金 >= dblSubBalance Then
        mFeeBalance.低保帐户支付 = dblSubBalance
    Else
        mFeeBalance.低保帐户支付 = dbl现金
    End If
    str结算方式 = str结算方式 & "|低保帐户;" & mFeeBalance.低保帐户支付 & ";0"
    
    gblnCancel_南京 = False '此时不允许取消结算
    门诊虚拟结算_南京市 = True
    Exit Function
errorhandle:
    Call DebugTool("虚拟结算时发生错误：" & Err.Description)
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_南京市(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    Dim intDO As Integer
    Dim lng病人ID As Long
    Dim intOrder As Integer
    Dim strNO As String
    Dim strRecord As String         '记录使用的发票清单
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim intSubInsure As Integer, strYHLB As String, dblSubBalance As Double, strSubInsureNO As String, intSubDisable As Integer  '子医保序号，子医保帐户余额及子医保号，停用标志
    Dim str交易号 As String, str入参1 As String, str入参2 As String, str出参 As String
    On Error GoTo errorhandle
    
    gstrSQL = " Select 病人ID From 门诊费用记录 Where 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人ID", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    
    '提取当前病人子医保相关信息(保险序号|优惠类别|医保号|余额|停用)
    gstrSQL = " Select 退休证号||'||||' AS 退休证号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前病人子医保相关信息", TYPE_南京市, lng病人ID)
    intSubInsure = Val(Split(rsTemp!退休证号, "|")(0))
    strYHLB = Split(rsTemp!退休证号, "|")(1)
    strSubInsureNO = Split(rsTemp!退休证号, "|")(2)
    dblSubBalance = Val(Split(rsTemp!退休证号, "|")(3))
    intSubDisable = Val(Split(rsTemp!退休证号, "|")(4))
    Select Case strYHLB
    Case "惠民"
        strYHLB = "1"
    Case "慈善"
        strYHLB = "2"
    Case "零差率"
        strYHLB = "3"
    Case Else
        strYHLB = "0"
    End Select
    '完成低保交易
    If intSubInsure <> 0 Then
        If Not CreateObject_Insure(intSubInsure, intOrder) Then Exit Function
        
        str交易号 = "05"
        str入参1 = strSubInsureNO & "|" & ToVarchar(gstr单位名称, 50) & "|" & lng结帐ID & "|" & mFeeBalance.门诊费用合计 & _
                  "|" & 0 & "|" & mFeeBalance.低保帐户支付 & "|" & UserInfo.姓名 & "|" & gstr备注
        If Not gobjInsure_Obj(intOrder).CallAPI(str交易号, str入参1, str入参2, str出参) Then Exit Function
    End If
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_南京市 & "," & mstrPatID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              mFeeBalance.门诊费用合计 & "," & mFeeBalance.自理费用 + mFeeBalance.个人自付 & ",0," & _
              mFeeBalance.医保范围费用 & "," & mFeeBalance.统筹支付 & "," & mFeeBalance.大病支付 & "," & _
              "0," & mFeeBalance.个人帐户支付 & ",'" & mFeeBalance.单据号 & "',null,null,'" & gPatInfo_南京市.病种编码 & "|" & gPatInfo_南京市.病种名称 & IIf(str出参 = "", "", "|" & Split(str出参, "|")(0) & "|" & intSubInsure) & "|" & mFeeBalance.医保卡号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "南京市医保")
    
    '保存票据使用记录
'    zl_票据使用明细_Insert (
'    领用ID_IN IN 票据使用明细.领用ID%TYPE,
'    票种_IN IN 票据使用明细.票种%TYPE,
'    号码_IN IN 票据使用明细.号码%TYPE,
'    性质_IN IN 票据使用明细.性质%TYPE,
'    原因_IN IN 票据使用明细.原因%TYPE,
'    结帐ID_IN IN 票据使用明细.结帐ID%TYPE,
'    使用时间_IN IN 票据使用明细.使用时间%TYPE,
'    使用人_IN IN 票据使用明细.使用人%TYPE
    If gblnBill Then
        gcnNJSYB.BeginTrans
        blnTrans = True
        For intDO = 1 To gintBills
            strNO = GetNextBill(glng领用ID)
            gstrSQL = "zl_票据使用明细_Insert(" & glng领用ID & ",1,'" & strNO & "',1,1," & lng结帐ID & ",sysdate,'" & UserInfo.姓名 & "')"
            gcnNJSYB.Execute gstrSQL, , adCmdStoredProc
            strRecord = strRecord & "," & strNO
        Next
        gcnNJSYB.CommitTrans
        blnTrans = False
        
        If strRecord <> "" Then
            strRecord = Mid(strRecord, 2)
            Err.Raise 9000, gstrSysName, "本次医保使用票据号：" & strRecord
        End If
    End If
    
    gblnCancel_南京 = True
    门诊结算_南京市 = True
    Exit Function
    
errorhandle:
    If blnTrans Then gcnNJSYB.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_南京市(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    Dim intOrder As Integer, intSubInsure As Integer, strSub顺序号 As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim lng销帐ID As Long
    Dim str交易号 As String, str入参1 As String, str入参2 As String, str出参 As String
    On Error GoTo errorhandle
    
    gstrSQL = "select distinct A.结帐id  from 门诊费用记录 A,门诊费用记录 B where A.记录状态=2 and A.NO=B.NO and B.结帐id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "销帐id", lng结帐ID)
    lng销帐ID = rsTemp!结帐ID
    
    gstrSQL = "select * from 保险结算记录 where 记录id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "原始记录", lng结帐ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "保险结算记录中原始结帐单据不存在,不允许退费"
        Exit Function
    Else
        gstrSQL = "zl_保险结算记录_insert(1," & lng销帐ID & "," & TYPE_南京市 & "," & rsTemp!病人ID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              -rsTemp!发生费用金额 & "," & -rsTemp!全自付金额 & "," & -rsTemp!首先自付金额 & "," & -rsTemp!进入统筹金额 & "," & -rsTemp!统筹报销金额 & "," & -rsTemp!大病自付金额 & "," & _
              "0," & -rsTemp!个人帐户支付 & ",null,null,null,null)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "销帐记录")
    End If
    
    If InStr(1, rsTemp!支付顺序号, "|") <> 0 Then
        intSubInsure = Val(Split(rsTemp!支付顺序号, "|")(2))
        strSub顺序号 = Split(rsTemp!支付顺序号, "|")(1)
        If Not CreateObject_Insure(intSubInsure, intOrder) Then Exit Function
        
        str交易号 = "06"
        str入参1 = strSub顺序号 & "|" & UserInfo.姓名
        If Not gobjInsure_Obj(intOrder).CallAPI(str交易号, str入参1, str入参2, str出参) Then Exit Function
    End If
    
    If gblnBill Then
        gcnNJSYB.BeginTrans
        blnTrans = True
        '产生作废收回废票记录
        gstrSQL = " Select * From 票据使用明细 Where 结帐ID=" & lng结帐ID
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open gstrSQL, gcnNJSYB
        With rsTemp
            Do While Not .EOF
                gstrSQL = "zl_票据使用明细_Insert(" & !领用ID & ",1,'" & !号码 & "',2,2," & lng销帐ID & ",sysdate,'" & UserInfo.姓名 & "')"
                gcnNJSYB.Execute gstrSQL, , adCmdStoredProc
                .MoveNext
            Loop
        End With
        gcnNJSYB.CommitTrans
        blnTrans = False
    End If
    
    门诊结算冲销_南京市 = True
    Exit Function
errorhandle:
    If blnTrans Then gcnNJSYB.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院虚拟结算_南京市(rsExse As Recordset, ByVal lng病人ID As Long, Optional strAdvance As String) As String
    Dim bytType As Byte
    Dim strFile As String, strWrite As String
    Dim strStream As String
    Dim dblSettleSum As Double
    Dim rsTemp As New ADODB.Recordset
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim a As Double
    '删除可能存在的前次结算信息文件
    On Error Resume Next
    
    strAdvance = ""
    If gblnBill Then
        glng公用ID = GetSetting("ZLSOFT", "公共模块\医保票据管理\住院", "共用收费票据批次", 0)
        glng领用ID = GetInvoiceGroupID(3, 1, glng领用ID, glng公用ID)
        If glng领用ID <= 0 Then
            Select Case glng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的医保收费票据,请先领用一批医保票据或设置本地共用医保票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地共用的医保票据已经用完,请先领用一批医保票据或重新设置本地共用医保票据！", vbInformation, gstrSysName
            End Select
            Exit Function
        End If
        strAdvance = GetNextBill(glng领用ID) '向前台程序返回本次结算所使用的票据开始号码
    End If
    
    Call Kill("C:\NJYB\CYJSD.XML")
    On Error GoTo errorhandle
    '上传还未上传的明细费用
    gstrSQL = "select 顺序号 from 保险帐户 where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "顺序号", lng病人ID)
    mDetailFee_南京市.住院序号 = rsTemp!顺序号
    
    '未对码项目
    gstrSQL = "select b.名称 as 项目 from 住院费用记录 a ,收费项目目录 b  where a.收费细目id=b.id and " & _
              "a.病人id=[1] and a.主页id= [2] and not exists ( select 1 from 保险支付项目 d where d.险类=[3] and a.收费细目id=d.收费细目id)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "收细明细", CLng(rsExse!病人ID), CLng(rsExse!主页ID), TYPE_南京市)
    If rsTemp.RecordCount <> 0 Then
        MsgBox "项目有未对编码: " & rsTemp!项目, vbInformation, gstrSysName
        Exit Function
    End If
    
    '打开文件
'    strFile = "C:\NJYB\ZYFYMX.XML"
'    Call writeTxtFile(strFile, "")
    If InitXML = False Then Exit Function
    Do Until rsExse.EOF
        If rsExse!是否上传 = 1 Or rsExse!金额 = 0 Then GoTo haddeliver           '找出已上传记录
        gstrSQL = "select Rownum 序号,decode(A.类别,'5',0,'6',0,'7',0,1) 标志,A.名称,A.编码,C.项目编码,A.计算单位,B.产地,decode(B.药品来源,'国产',1,'合资',2,'进口',3,null) 产地特征,B.规格" & _
                  " from 收费细目 A,药品目录 B,保险支付项目 C where A.id = C.收费细目id and A.id=B.药品id(+) and A.id = [1] And C.险类=[2]"
'        gstrSQL = "select Rownum 序号,decode(A.类别,'5',0,'6',0,'7',0,1) 标志,A.名称,A.编码,C.项目编码,A.计算单位,B.产地,decode(B.药品来源,'国产',1,'合资',2,'进口',3,null) 产地特征,B.规格,d.现价" & _
'                  " from 收费细目 A,药品目录 B,保险支付项目 C,收费价目 d where A.id = C.收费细目id and a.id=d.id and A.id=B.药品id(+) and A.id =" & rsExse!收费细目ID & " And C.险类=" & TYPE_南京市
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "收细明细", CLng(rsExse!收费细目ID), TYPE_南京市)
        If rsTemp.RecordCount <> 0 Then
            With mDetailFee_南京市
                .行号 = rsTemp!序号
                .标志 = rsTemp!标志
                .费用发生时间 = Format(rsExse!登记时间, "yyyyMMdd")
                .医院自编码 = rsTemp!编码
                .医保编码 = rsTemp!项目编码
                .名称 = rsTemp!名称
                .剂量单位 = zlCommFun.Nvl(rsTemp!计算单位)
                .单价 = Format(rsExse!金额 / rsExse!数量, "#0.0000;-#0.0000;0;")
'                .单价 = Format(rsTemp!现价, "#0.0000;-#0.0000;0;")
                .数量 = rsExse!数量
                .产地 = zlCommFun.Nvl(rsTemp!产地)
                .产地特征 = zlCommFun.Nvl(rsTemp!产地特征)
                .规格 = zlCommFun.Nvl(rsTemp!规格)
            End With
            
            gstrSQL = "select b.电子邮件 as 医生编码,a.开单人 as 医生 from 住院费用记录 a,人员表 b where a.NO=[1] and a.序号=[2]" & _
                    " and mod(a.记录性质,10)=[3] and a.记录状态=[4] And a.开单人 = b.姓名"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医生编码", CStr(rsExse!NO), CLng(rsExse!序号), CInt(rsExse!记录性质), CInt(rsExse!记录状态))
            If rsTemp.RecordCount = 0 Then
                '取住院医师
                gstrSQL = "select b.电子邮件 as 医生编码,a.住院医师 as 医生 from 病案主页 a,人员表 b  where a.住院医师=b.姓名 and  " & _
                          "a.病人id= [1] and a.主页id= [2] and rownum<2"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取住院医师", CLng(rsExse!病人ID), CLng(rsExse!主页ID))
            End If
            mDetailFee_南京市.操作人编码 = rsTemp!医生编码
            mDetailFee_南京市.操作人姓名 = rsTemp!医生
            
            a = a + 1
            Set nodRow = InsertChild(mdomInput.documentElement, "RECORD", "")
            Call InsertChild(nodRow, "ID", a)
            Call InsertChild(nodRow, "XH", mDetailFee_南京市.住院序号)
            Call InsertChild(nodRow, "BZ", mDetailFee_南京市.标志)
            Call InsertChild(nodRow, "SJ", mDetailFee_南京市.费用发生时间)
            Call InsertChild(nodRow, "ZBM", mDetailFee_南京市.医保编码)
            Call InsertChild(nodRow, "SL", mDetailFee_南京市.数量)
            Call InsertChild(nodRow, "DJ", mDetailFee_南京市.单价)
            Call InsertChild(nodRow, "YSM", mDetailFee_南京市.操作人编码)
            Call InsertChild(nodRow, "YS", mDetailFee_南京市.操作人姓名)
        End If
haddeliver:
        dblSettleSum = dblSettleSum + rsExse!金额           '得出结帐总金额
        rsExse.MoveNext
    Loop
    '关闭文件
    mdomInput.Save "C:\NJYB\zyfymx.xml"
'    Call writeTxtFile(strFile, "", False)
    
    bytType = 9                          '表示住院预结算状态
    
    strStream = frm数据交换.getFeeBalance(bytType)
    On Error Resume Next
    Unload frm数据交换
    On Error GoTo errorhandle
    If strStream = "" Then
        MsgBox "读取医保结算文件过程被中止,无法完成预结算", vbInformation, gstrSysName
        Exit Function
    End If
    
'    With mFeeBalance
'        .住院序号 = analyseStr(strStream, 1, 20)
'        .门诊费用合计 = Val(analyseStr(strStream, 35, 10))
'        .医保范围费用 = Val(analyseStr(strStream, 65, 10))
'        .自理费用 = Val(analyseStr(strStream, 75, 10))
'        .个人自付 = Val(analyseStr(strStream, 85, 10))
'        .统筹支付 = Val(analyseStr(strStream, 95, 10))
'        .大病支付 = Val(analyseStr(strStream, 105, 10))
'        .个人帐户支付 = Val(analyseStr(strStream, 115, 10))
'    End With
    Set mdomInput = New MSXML2.DOMDocument
    If mdomInput.Load("c:\njyb\cyjsd.xml") = False Then
        MsgBox "医保服务器返回值格式不正确。", vbInformation, gstrSysName
    Else
        Set nodRowset = mdomInput.documentElement.selectSingleNode("RECORD")
          With mFeeBalance
         .住院序号 = nodRowset.selectSingleNode("XH").Text
         .门诊费用合计 = nodRowset.selectSingleNode("ZFY").Text
         .自理费用 = nodRowset.selectSingleNode("GRZL").Text
         .个人自付 = nodRowset.selectSingleNode("GRZF").Text
         .统筹支付 = nodRowset.selectSingleNode("YBZF").Text
         .个人帐户支付 = nodRowset.selectSingleNode("ZHZF").Text
         End With
    End If
    mcur个帐余额 = nodRowset.selectSingleNode("ZHYE").Text
    If mFeeBalance.住院序号 <> mDetailFee_南京市.住院序号 Then
        MsgBox "此结帐病人与医保结算文件中病人不一致,不能结算", vbInformation, gstrSysName
        Exit Function
    End If
    If Format(dblSettleSum, "#0.00") <> Format(mFeeBalance.门诊费用合计, "#0.00") Then
        MsgBox "请注意:医院总费用与医保中心返回的总费用不一致" & vbCrLf & _
        "总费用:(医院)￥" & Format(dblSettleSum, "#0.00") & Space(10) & "(医保)￥" & Format(mFeeBalance.门诊费用合计, "#0.00"), vbInformation, gstrSysName
    End If

    strStream = "统筹基金;" & mFeeBalance.统筹支付 & ";0"
    If mFeeBalance.个人帐户支付 <> 0 Then
        strStream = strStream & "|个人帐户;" & mFeeBalance.个人帐户支付 & ";0"
    End If
    If mFeeBalance.大病支付 <> 0 Then
        strStream = strStream & "|大病统筹;" & mFeeBalance.大病支付 & ";0"
    End If
    
    住院虚拟结算_南京市 = strStream
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_南京市(lng结帐ID As Long, lng病人ID) As Boolean
    Dim strNO As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errorhandle
    gstrSQL = "select NO,序号,记录状态,记录性质 from 住院费用记录 where nvl(是否上传,0)=0 and 结帐id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查找记录", lng结帐ID)
    Do Until rsTemp.EOF
        gstrSQL = "ZL_病人费用记录_上传('" & rsTemp!NO & "'," & rsTemp!序号 & "," & rsTemp!记录性质 & "," & rsTemp!记录状态 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新上传标志")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "select 住院次数 from 病人信息 where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "主页id", lng病人ID)
    
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_南京市 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              mFeeBalance.门诊费用合计 & "," & mFeeBalance.自理费用 + mFeeBalance.个人自付 & ",0," & _
              mFeeBalance.医保范围费用 & "," & mFeeBalance.统筹支付 & "," & mFeeBalance.大病支付 & "," & _
              "0," & mFeeBalance.个人帐户支付 & ",'" & mFeeBalance.住院序号 & "'," & rsTemp!住院次数 & ",null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "插入保险帐户")
    
    '保存票据使用记录（住院只可能是一张票据）
'    zl_票据使用明细_Insert (
'    领用ID_IN IN 票据使用明细.领用ID%TYPE,
'    票种_IN IN 票据使用明细.票种%TYPE,
'    号码_IN IN 票据使用明细.号码%TYPE,
'    性质_IN IN 票据使用明细.性质%TYPE,
'    原因_IN IN 票据使用明细.原因%TYPE,
'    结帐ID_IN IN 票据使用明细.结帐ID%TYPE,
'    使用时间_IN IN 票据使用明细.使用时间%TYPE,
'    使用人_IN IN 票据使用明细.使用人%TYPE
    If gblnBill Then
        strNO = GetNextBill(glng领用ID)
        gstrSQL = "zl_票据使用明细_Insert(" & glng领用ID & ",3,'" & strNO & "',1,1," & lng结帐ID & ",sysdate,'" & UserInfo.姓名 & "')"
        gcnNJSYB.Execute gstrSQL, , adCmdStoredProc
        Err.Raise 9000, gstrSysName, "本次医保使用票据号：" & strNO
    End If
    
    住院结算_南京市 = True
    Exit Function
errorhandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_南京市(lng结帐ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim lng销帐ID As Long
    
    On Error GoTo errorhandle
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "销帐id", lng结帐ID)
    lng销帐ID = rsTemp!ID
    
    gstrSQL = "select * from 保险结算记录 where 记录id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "原始记录", lng结帐ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "保险结算记录中原始结帐单据不存在,不允许退费"
        Exit Function
    Else
        gstrSQL = "zl_保险结算记录_insert(2," & lng销帐ID & "," & TYPE_南京市 & "," & rsTemp!病人ID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              -rsTemp!发生费用金额 & "," & -rsTemp!全自付金额 & "," & -rsTemp!首先自付金额 & "," & -rsTemp!进入统筹金额 & "," & -rsTemp!统筹报销金额 & "," & -rsTemp!大病自付金额 & "," & _
              "0," & -rsTemp!个人帐户支付 & ",'" & rsTemp!支付顺序号 & "'," & rsTemp!主页ID & ",null,null)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "销帐记录")
    End If
    
    If gblnBill Then
        gcnNJSYB.BeginTrans
        blnTrans = True
        '产生作废收回废票记录
        gstrSQL = " Select * From 票据使用明细 Where 结帐ID=" & lng结帐ID
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open gstrSQL, gcnNJSYB
        With rsTemp
            Do While Not .EOF
                gstrSQL = "zl_票据使用明细_Insert(" & !领用ID & ",3,'" & !号码 & "',2,2," & lng销帐ID & ",sysdate,'" & UserInfo.姓名 & "')"
                gcnNJSYB.Execute gstrSQL, , adCmdStoredProc
                .MoveNext
            Loop
        End With
        gcnNJSYB.CommitTrans
        blnTrans = False
    End If
    住院结算冲销_南京市 = True
    Exit Function
errorhandle:
    If blnTrans Then gcnNJSYB.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Sub writeTxtFile(strFile As String, strWrite As String, Optional ByVal openFile As Boolean = True)
    Dim intSymbol As Long
    Dim strFolder As String
    
    On Error GoTo errorhandle
    Do Until InStr(intSymbol + 1, strFile, "\") = 0
        intSymbol = InStr(intSymbol + 1, strFile, "\")
        strFolder = Mid(strFile, 1, intSymbol)
        If Not mobjSystem.FolderExists(strFolder) Then mobjSystem.CreateFolder (strFolder)
    Loop

    If openFile Then                    '打开文件
        If Not mobjSystem.FileExists(strFile) Then mobjSystem.CreateTextFile (strFile)
        Set mobjStream = mobjSystem.OpenTextFile(strFile, ForWriting)
        If strWrite <> "" Then          '如果有内容进行写入
            mobjStream.WriteLine (strWrite)
            mobjStream.Close
        End If
    Else
        If strWrite = "" Then
            mobjStream.Close
        Else
            mobjStream.WriteLine (strWrite)   '如果有写入内容但打开标志为false,只进行写入
        End If
    End If
    Exit Sub
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mobjStream.Close
End Sub

Public Function readTxtFile(strFile As String) As String
    On Error GoTo errHandle
    
    If mobjSystem.FileExists(strFile) Then
        Set mobjStream = mobjSystem.OpenTextFile(strFile)
        readTxtFile = mobjStream.ReadLine
        mobjStream.Close
    End If
    Exit Function
    
errHandle:
    Err.Clear
    On Error Resume Next
    mobjStream.Close
End Function

Private Function fillSpa(strTemp As Variant, lngLen As Long, Optional fromRigth As Boolean = True) As String
    Dim lngStrLeng As Long
    Dim strStream As String
    Dim strUnion As String
    
    strTemp = IIf(IsNull(strTemp), "", Trim(strTemp))
    
    strUnion = StrConv(Trim(strTemp), vbFromUnicode)
    lngStrLeng = IIf(LenB(strUnion) > lngLen, lngLen, LenB(strUnion))
    strStream = IIf(LenB(strUnion) > lngLen, StrConv(LeftB(strUnion, 20), vbUnicode), strTemp)
    
    If fromRigth Then
        fillSpa = strStream & String(lngLen - lngStrLeng, " ")
    Else
        fillSpa = String(lngLen - lngStrLeng, " ") & strStream
    End If
End Function

Public Function analyseStr(strTemp As String, lngStart As Long, lngLen As Long) As String
    Dim strStream As String
    
    strStream = StrConv(UCase(strTemp), vbFromUnicode)
    
    analyseStr = Trim(StrConv(MidB(strStream, lngStart, lngLen), vbUnicode))
End Function

Public Function 个人余额_南京市(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
'    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID='" & lng病人ID & "' and 险类=" & TYPE_南京市
'    Call OpenRecordset(rsTemp, gstrSysName)
'
'    If rsTemp.EOF Then
'        个人余额_南京市 = 100000
'    Else
'        个人余额_南京市 = IIf(rsTemp("帐户余额") = 0, 100000, rsTemp("帐户余额"))
'    End If
    个人余额_南京市 = 100000
End Function

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function GetFullNO(strNO As String) As String
'功能：由用户输入的部份单号，返回当年的单号。
    If Len(strNO) >= 8 Then GetFullNO = Right(strNO, 8): Exit Function
    GetFullNO = PreFixNO & Format(strNO, "0000000")
End Function

Public Function FileExists(ByVal FileName As String, Optional ErrFlag As Boolean = True) As Boolean
    Dim Temp
    FileExists = True
    On Error Resume Next
proshow:
    Temp = FileDateTime(FileName)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                If ErrFlag Then
                    If MsgBox("磁盘没有准备好。", vbInformation + vbRetryCancel, "错误") = vbRetry Then
                        GoTo proshow:
                    End If
                End If
                FileExists = False
            End If
    End Select
End Function


Public Function 挂号结算_南京市(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
'******************************************************************************
'调用者　　　　　　：该方法由门诊挂号部件调用
'功能说明　　　　　：通过调用医保商的门诊挂号接口，分解本次费用明细，得到结算结果（个人帐户多少、医保基金多少等）并保存
'注意事项　　　　　：如果存在个人帐户或医保基金或公务员补助，需要调用过程zl_病人结算记录_Update对病人预交记录进行数据修正
'调用过程清单及说明：
'　　【　　　】
''*****************************************************************************
    挂号结算_南京市 = True
End Function


Public Function 挂号结算冲销_南京市(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
'******************************************************************************
'调用者　　　　　　：该方法由门诊挂号部件调用
'功能说明　　　　　：通过调用医保商的门诊挂号冲销接口，完成门诊挂号结算的作废
'调用过程清单及说明：
'　　【　　　】
''*****************************************************************************
    挂号结算冲销_南京市 = True
End Function


Public Function 撤消医保入院_南京市(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str顺序号 As String) As Boolean
'功能：更新病人的出院疾病。如果是肿瘤，则结算时起付线会减半
    Dim StrInput As String, arrOutput  As Variant
    
    On Error GoTo errHandle
    
    gstrSQL = "ZL_病案主页_撤消医保入院(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "撤消医保入院")
    
    撤消医保入院_南京市 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function AnalyServer(ByVal strConn As String) As String
    Dim arrData, arrColumn
    Dim intDO As Integer, intMAX As Integer
    
    strConn = UCase(Replace(strConn, """", ""))
    arrData = Split(strConn, ";")
    intMAX = UBound(arrData)
    For intDO = 0 To intMAX
        arrColumn = Split(arrData(intDO), "=")
        If arrColumn(0) = "SERVER" Then
            AnalyServer = arrColumn(1)
            Exit Function
        End If
    Next
End Function

Private Sub AnalyConf(strUser As String, strPass As String, strServer As String)
    Dim arrLine
    Dim strLine As String
    Dim strFile As String
    Dim blnOpen As Boolean
    Dim objFileSys As New FileSystemObject
    Dim objStream As TextStream
    On Error GoTo errHand
    
    '从配置文件中读取医保前置机的用户名，密码与主机串
    strFile = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Conf.ini"
    If objFileSys.FileExists(strFile) Then
        Set objStream = objFileSys.OpenTextFile(strFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            If strLine = "" Then Exit Do
            arrLine = Split(strLine, "=")
            Select Case arrLine(0)
            Case "USER"
                strUser = arrLine(1)
            Case "PASS"
                strPass = arrLine(1)
            Case "SERVER"
                strServer = arrLine(1)
            End Select
        Loop
        objStream.Close
        blnOpen = False
    End If
    
    If strUser = "" Then strUser = "zl9I_NJSYB"
    If strPass = "" Then strPass = "HIS"
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub

Private Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intnum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, Optional ByVal strBill As String) As Long
'功能：获取张数够用并且指定票据在其可用范围内的领用ID
'参数：bytKind      =   票种
'      intNum       =   要打印的票据张数
'      lngLastUseID =   上次使用的领用ID
'      lngShareUseID=   本地参数指定的共用ID
'      strBill      =   当前票据号，用于检查领用批次的票据范围
'返回：
'      >0   =   成功，可用的领用ID
'      =0   =   失败
'      -1   =   没有自用(用完或不够，或未领用),未设置共用
'      -2   =   没有自用(用完或不够，或未领用),设置的共用已用完或不够
'      -3   =   指定票据号不在当前所有可用领用批次的有效票据号范围内
'      -4   =   指定批次的票据不够用
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    
    On Error GoTo ErrH
    '1.上次的领用批次是否可用并够用
    If lngLastUseID > 0 Then
        strSQL = "Select 前缀文本,开始号码,终止号码" & vbNewLine & _
                 "From 票据领用记录 Where 票种=" & bytKind & " And nvl(当前号码,开始号码)<>终止号码 And ID=" & lngLastUseID
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSQL, gcnNJSYB
        With rsTmp
            If .RecordCount > 0 Then    '目前的票据号可能和上次不同，所以需要检查范围
                If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '可能没有当前票据号
                blnTmp = False
                strPre = "" & !前缀文本
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
                
            ElseIf intnum > 1 Then  '不是确定领用批次调用时,当前票据号所在批次不够用
                GetInvoiceGroupID = -4: Exit Function
            End If
        End With
    End If
        
    '2.上次的领用批次不够用或不可用时,取自已领的并且自用的
    '  有多批，最近使用的优先,少的先用,先领先用
    strSQL = "Select ID, 前缀文本, 开始号码, 终止号码" & vbNewLine & _
        "From 票据领用记录" & vbNewLine & _
        "Where 票种 = " & bytKind & " And nvl(当前号码,开始号码)<>终止号码 And 领用人 = '" & UserInfo.姓名 & "' And 使用方式 = 1" & vbNewLine & _
        "Order By 开始号码"
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnNJSYB
    With rsTmp
        For i = 1 To .RecordCount
            If strBill = "" Then GetInvoiceGroupID = !ID: Exit Function '第一次使用时没有当前票据号
            blnTmp = False
            strPre = "" & !前缀文本
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = !ID: Exit Function
            .MoveNext
        Next
        lngReturn = IIf(.RecordCount > 0, -3, -1)
    End With
        
    '3.没有自用的,使用本地参数指定的共用批次
    If lngShareUseID > 0 Then
        strSQL = "Select 前缀文本,开始号码,终止号码" & vbNewLine & _
                 "From 票据领用记录 Where 票种=" & bytKind & " And nvl(当前号码,开始号码)<>终止号码 And ID=" & lngShareUseID
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSQL, gcnNJSYB
        With rsTmp
            If .RecordCount > 0 Then
                If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '第一次使用时没有当前票据号
                blnTmp = False
                strPre = "" & !前缀文本
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
            End If
            lngReturn = IIf(.RecordCount > 0, -3, -2)
        End With
    End If
    
    GetInvoiceGroupID = lngReturn   '返回未找到的原因代码
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckUsedBill(bytKind As Byte, ByVal lng领用ID As Long, Optional ByVal strBill As String) As Long
'功能：检查当前操作员是否有可用票据领用(自用或共用),并返回可用的领用ID
'参数：bytKind=票种
'      lng领用ID=第一次检查时为本地设置的共用领用ID,以后为上次使用的领用ID
'      strBill=要检查范围的票据号
'说明：
'    1.在检查范围时,如果病人有多批自用票据,则只要在其中一批之中就行了
'    2.在检查范围时,长度也在检查范围之内。
'    3.当有多批自用时,缺省按少的先用,先领先用,"最近使用的优先"原则
'返回：
'      正常：票据领用ID>0
'      0=失败
'      -1:没有自用(用完或未领用)、也没有共用(未设置)
'      -2:设置的共用已用完
'      -3:指定票据号不在当前可用范围内(包含多批自用票据的情况)

    Dim rsTmp As New ADODB.Recordset
    Dim rsSelf As New ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo ErrH
    
    '操作员有剩余的自用票据集
    strSQL = _
        "Select ID, 前缀文本, 开始号码, 终止号码, 登记时间, 使用时间" & vbNewLine & _
        "From 票据领用记录" & vbNewLine & _
        "Where 票种 = " & bytKind & " And 使用方式 = 1 And nvl(当前号码,开始号码)<>终止号码 And 领用人 = '" & UserInfo.姓名 & "'" & vbNewLine & _
        "Order By 开始号码"
    If rsSelf.State = 1 Then rsSelf.Close
    rsSelf.CursorLocation = adUseClient
    rsSelf.Open strSQL, gcnNJSYB
    
    If lng领用ID = 0 Then
        '程序中第一次检查,且没有设置本地共用
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '也没有自用票据
        '有自用票据,按优先原则返回
        lngReturn = rsSelf!ID
    Else
        '上次使用的领用ID或第一次检查的共用ID,先判断性质
        strSQL = "Select ID,使用方式,前缀文本,开始号码,终止号码 From 票据领用记录 Where nvl(当前号码,开始号码)<>终止号码 And 票种=" & bytKind & " And ID=" & lng领用ID
        If rsSelf.State = 1 Then rsSelf.Close
        rsSelf.CursorLocation = adUseClient
        rsSelf.Open strSQL, gcnNJSYB
        If rsTmp!使用方式 = 2 Then '共用,要先看有没有自用
            If Not rsSelf.EOF Then
                '有自用的，优先
                lngReturn = rsSelf!ID
            Else
                '没有自用取共用
                'If rsTmp!剩余数量 = 0 Then CheckUsedBill = -2: Exit Function '共用已经用完
                lngReturn = rsTmp!ID
                blnTmp = True
            End If
        Else
            '自用票据
'            If rsTmp!剩余数量 > 0 Then
                '有剩余
                lngReturn = rsTmp!ID
'            Else
'                '其它有剩余的自用
'                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '其它自用也没有剩余
'                lngReturn = rsSelf!ID
'            End If
        End If
    End If
    
    '检查票号范围是否正确
    If strBill <> "" Then
        If blnTmp Then
            '在共用范围内范围判断
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)))) <> UCase(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!开始号码) And UCase(strBill) <= UCase(rsTmp!终止号码) And Len(strBill) = Len(rsTmp!开始号码)) Then
                lngReturn = -3
            End If
        Else
            '在可用自用范围内判断
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)))) <> UCase(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!开始号码) And UCase(strBill) <= UCase(rsSelf!终止号码) And Len(strBill) = Len(rsSelf!开始号码)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '该批不满足,则在其它自用中检查
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)))) <> UCase(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!开始号码) And UCase(strBill) <= UCase(rsSelf!终止号码) And Len(strBill) = Len(rsSelf!开始号码)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CheckUsedBill = 0
End Function

Private Function GetNextBill(lng领用ID As Long) As String
'功能：根据领用批次ID,获取下一个实际票据号
'说明：1.当取不到范围内的有效票据时,返回空由用户输入
'      2.排开已报损的号码
    Dim rsMain As New ADODB.Recordset
    Dim rsDelete As New ADODB.Recordset
    Dim strSQL As String, strBill As String
    
    On Error GoTo ErrH
    
    strSQL = "Select 前缀文本,开始号码,终止号码,当前号码" & _
        " From 票据领用记录 Where nvl(当前号码,开始号码)<>终止号码 And ID=" & lng领用ID
    If rsMain.State = 1 Then rsMain.Close
    rsMain.CursorLocation = adUseClient
    rsMain.Open strSQL, gcnNJSYB
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!当前号码) Then
        strBill = UCase(rsMain!开始号码)
    Else
        strBill = UCase(IncStr(rsMain!当前号码))
    End If
    
    strSQL = "Select Upper(号码) as 号码 From 票据使用明细" & _
        " Where 性质=1 And 原因=5 And 号码>='" & strBill & "' And 领用ID=" & lng领用ID & _
        " Order by 号码"
    If rsDelete.State = 1 Then rsDelete.Close
    rsDelete.CursorLocation = adUseClient
    rsDelete.Open strSQL, gcnNJSYB
    Do While True
        '检查范围
        If Left(strBill, Len("" & rsMain!前缀文本)) <> UCase("" & rsMain!前缀文本) Then
            Exit Function
        ElseIf Not (strBill >= UCase(rsMain!开始号码) And strBill <= UCase(rsMain!终止号码)) Then
            Exit Function
        End If
                
        '排开报损号
        rsDelete.Filter = "号码='" & UCase(strBill) & "'"
        If rsDelete.EOF Then Exit Do
        strBill = IncStr(strBill)
    Loop
   
    GetNextBill = strBill
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Private Function AnalyBill(ByVal rsDetail As ADODB.Recordset) As Long
    Dim intBills_A As Integer, intBills_B As Integer
    Dim int明细数 As Integer, int收据费目 As Integer
    Dim str收据费目 As String
    Dim rs票据费目 As New ADODB.Recordset
    '返回本次结算的票据张数
    
    With rsDetail
        int明细数 = .RecordCount
        Do While Not .EOF
            On Error Resume Next
            Err = 0
            '取打印时应该使用的收据费目（V10。18才增加这张表）
            gstrSQL = "select 收据费目 from 收据费目对应 a,收费价目 b where a.场合=0 and a.收入项目id=b.收入项目id and (b.终止日期>sysdate or b.终止日期=to_date('3000-01-01','yyyy-mm-dd') )" & _
                     " and 执行日期<sysdate and b.收费细目id= [1]"
            Set rs票据费目 = zlDatabase.OpenSQLRecord(gstrSQL, "票据费目", CLng(!收费细目ID))
            If Err = 0 Then
                If InStr(1, str收据费目, rs票据费目!收据费目) = 0 Then
                    str收据费目 = str收据费目 & "," & rs票据费目!收据费目
                    int收据费目 = int收据费目 + 1
                End If
            Else
                If InStr(1, str收据费目, !收据费目) = 0 Then
                    str收据费目 = str收据费目 & "," & !收据费目
                    int收据费目 = int收据费目 + 1
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    '计算本次票据使用张数,取最大
    intBills_A = int明细数 \ gint明细数
    If int明细数 Mod gint明细数 <> 0 Then intBills_A = intBills_A + 1
    intBills_B = int收据费目 \ gint收据费目
    If int收据费目 Mod gint收据费目 <> 0 Then intBills_B = intBills_B + 1
    
    If intBills_A >= intBills_B Then
        AnalyBill = intBills_A
    Else
        AnalyBill = intBills_B
    End If
End Function

'检查剩下的票据张数是否够用，不够用则提示并不允许进行结算
Private Function IsEnough() As Boolean
    Dim lng当前号码 As Long, lng终止号码 As Long
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = " Select 前缀文本,终止号码 From 票据领用记录 Where ID=" & glng领用ID
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnNJSYB
    lng当前号码 = Mid(GetNextBill(glng领用ID), Len(rsTemp!前缀文本) + 1)
    lng终止号码 = Mid(rsTemp!终止号码, Len(rsTemp!前缀文本) + 1)
    IsEnough = (lng终止号码 - lng当前号码 + 1 >= gintBills)
End Function

Private Function IncStr(ByVal strVal As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
    Dim i As Integer, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function
