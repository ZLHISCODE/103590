Attribute VB_Name = "mdl神木大兴"
Option Explicit
Private mblnInit As Boolean     '是否已经初始化

Private Type InitbaseInfor
    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    医院编码 As String                      '初始医院编码
    病人目录 As String
    
End Type


Public InitInfor_神木大兴 As InitbaseInfor
Private Type 病人身份
    IC卡号              As String
    姓名                As String
    性别                As String
    
    结帐ID              As Long         '当前结帐ID值
    病人ID              As Long
    当前门诊号          As String       'right(病人ID1,6)-yyyyMMDDHHMMSS
    费用总额            As Double
    
    虚拟结算            As Boolean  '已经虚拟结算,或结算
End Type


Private Type 结算数据
    费用总额          As Double       '门诊和住院
    个人帐户支付    As Double       '门诊和住院
    统筹支付        As Double       '门诊和住院
    公务员床补      As Double       '住院
    押金总额        As Double       '住院
    应补现金额      As Double       '住院
    公费床位费      As Double       '住院
    自费床位费      As Double       '住院
    公费调温费      As Double       '住院
    自费调温费      As Double       '住院
    结算前卡上余额  As Double       '门诊
    结算后卡上余额  As Double       '门诊
End Type
Private g结算数据 As 结算数据

Public g病人身份_神木大兴 As 病人身份
Public gcnOracle_神木大兴 As ADODB.Connection     '中间库连接

Public Function 医保初始化_神木大兴() As Boolean
    
    Dim strReg As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim strUser As String, strPass As String, strServer As String
    
    If mblnInit = True Then
        医保初始化_神木大兴 = True
        Exit Function
    End If
    
    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_神木大兴.模拟数据 = True
    Else
        InitInfor_神木大兴.模拟数据 = False
    End If
   
    InitInfor_神木大兴.医院编码 = gstr医院编码
    InitInfor_神木大兴.病人目录 = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("病人目录"), App.Path) & "\ReadYbInfo.INI"
    
    
    If Open中间库_神木大兴 = False Then
        Exit Function
    End If
    mblnInit = True
    医保初始化_神木大兴 = True
End Function

Public Function 医保终止_神木大兴() As Boolean
    
    '将初始化标志置为false
    mblnInit = False
    If gcnOracle_神木大兴.State = 1 Then
        gcnOracle_神木大兴.Close
    End If
    医保终止_神木大兴 = True
End Function

Public Function 身份标识_神木大兴(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '返回：空或信息串
    Err = 0
    On Error GoTo errHand:
    身份标识_神木大兴 = frmIdentify神木大兴.GetPatient(bytType, lng病人ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_神木大兴 = ""
End Function


Public Function 个人余额_神木大兴(ByVal lng病人ID As Long, ByRef dbl透支额 As Currency) As Currency
    '功能: 提取参保病人个人帐户余额
    '返回: 返回个人帐户余额
    dbl透支额 = 10000000000000#
    个人余额_神木大兴 = 0
End Function
Private Function 获取个人帐户支付() As Double
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取个人帐户值(从预交记录中获取)
    '--入参数:
    '--出参数:
    '--返  回:成功,返回本次个人帐户支付,否则返回0
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select * From 病人预交记录 where 结帐ID=[1] and  结算方式='个人帐户'"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取个人帐户支付", g病人身份_神木大兴.结帐ID)
    If Not rsTemp.EOF Then
        获取个人帐户支付 = Nvl(rsTemp!冲预交, 0)
    End If
End Function

Private Function Check医保项目(ByVal lng收费细目ID As Long, ByRef str类别 As String, ByRef str医保编码 As String, ByRef str医保名称 As String, ByRef str拼音编码 As String, Optional bln忽略 As Boolean = False) As Boolean
    '功能:获取相关的医保项目信息
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "" & _
              "   Select a.附注,a.项目编码,a.项目名称,b.编码,b.名称 " & _
              "   From 保险支付项目 a,收费细目 B  " & _
              "   where a.收费细目id=b.ID and  a.险类=" & TYPE_陕西大兴 & _
              "           and a.收费细目id=" & lng收费细目ID
    
    Check医保项目 = False
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取保险项目"
    If rsTemp.EOF Then
        If bln忽略 = False Then
                gstrSQL = "Select 编码,名称 From 收费细目 where ID=" & lng收费细目ID
                zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取保险项目"
                ShowMsgbox "收费项目“" & Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称) & "”还未进行医保对码，不能继续操作!"
        End If
        Exit Function
    End If
    str类别 = Nvl(rsTemp!附注)
    str医保编码 = Nvl(rsTemp!项目编码)
    str医保名称 = Nvl(rsTemp!项目名称)

    gstrSQL = "select pybm from yy_ypfzb  where lb='" & str类别 & "' and bm='" & str医保编码 & "' and mc='" & str医保名称 & "'"
    
    Call OpenRecordset_神木大兴(rsTemp, "获取拼音码", gstrSQL)
    If Not rsTemp.EOF Then
        str拼音编码 = Nvl(rsTemp!pybm)
    Else
        str拼音编码 = ""
    End If
    Check医保项目 = True
End Function
Public Function 门诊虚拟结算取消_神木大兴(ByVal bytType As Byte, ByVal lng病人ID As Long) As Boolean
   '-----------------------------------------------------------------------------------------------------------
    '--功  能:按取消按钮
    '--入参数:
    '--出参数:
    '--返  回:成功,true
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim blnYes As Boolean
    
    门诊虚拟结算取消_神木大兴 = False
    If g病人身份_神木大兴.虚拟结算 = False Then Exit Function
    
    gstrSQL = "Select YBKH,CFBH,YXBZ From MZ_FYMXB where YBKH='" & g病人身份_神木大兴.IC卡号 & "' and rownum<=1 "
    Call OpenRecordset_神木大兴(rsTemp, "检查是否已经结算过", gstrSQL)
    If rsTemp.RecordCount <> 0 Then
        ShowMsgbox "该病人医保已经结算过了,但是你人为地取消了该操作," & vbCrLf & "这样会造成中心与医院数据不符,真的要强制退出吗?", True, blnYes
        If Not blnYes Then
            Exit Function
        End If
    End If
    门诊虚拟结算取消_神木大兴 = True

End Function
Public Function 门诊虚拟结算_神木大兴(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保

    Dim str明细 As String
    Dim rsTemp As New ADODB.Recordset
    Dim str类别 As String, str医保编码 As String, str医保名称 As String, str拼音码 As String
    
    
    g病人身份_神木大兴.费用总额 = 0
    
    If g病人身份_神木大兴.虚拟结算 = True Then
        ShowMsgbox "已经虚拟结算过了,请按结算按钮!"
        Exit Function
    End If
    
    str明细 = ""
    If rs明细.RecordCount <> 0 Then
        g病人身份_神木大兴.当前门诊号 = Lpad(Right(CStr(rs明细!病人ID), 6), 6, "0") & Format(zlDatabase.Currentdate, "yyyymmddHHMMSS") 'right(病人ID1,6)-yyyyMMDDHHMMSS
    Else
        g病人身份_神木大兴.当前门诊号 = ""
    End If
    
    
    '第一步:判断是否存在未结算的明细费用
    
    Err = 0: On Error GoTo errHand:
    gcnOracle_神木大兴.BeginTrans
    
    DebugTool "门诊虚拟结算,第一步:判断是否存在未结算的明细费用"
    
    gstrSQL = "Select YBKH,CFBH,YXBZ From MZ_FYMXB where YBKH='" & g病人身份_神木大兴.IC卡号 & "' and rownum<=1 "
    Call OpenRecordset_神木大兴(rsTemp, "检查是否已经结算过", gstrSQL)
    If rsTemp.RecordCount <> 0 Then
        gstrSQL = "Select YBKH,CFBH,YXBZ From MZ_FYMXB where YBKH='" & g病人身份_神木大兴.IC卡号 & "' and YXBZ<>'F' and rownum<=1 "
        Call OpenRecordset_神木大兴(rsTemp, "检查是否已经结算过", gstrSQL)
        If Not rsTemp.EOF Then
            Dim blnYes As Boolean
            ShowMsgbox "该医保病人上次已经提交了明细,但未清除," & vbCrLf & "(处方号" & Nvl(rsTemp!cfbh) & "),可能原因如下:" & vbCrLf & "    1.可能操作员在结算过程中中途终止了!" & vbCrLf & "    2.可能在上传明细完成后,程序出现非正式退出!" & vbCrLf & "    3.可能已经虚拟结算了,但HIS还未正式结算!" & vbCrLf & " 是否要进行强制清楚?", True, blnYes
            If blnYes = False Then
                gcnOracle_神木大兴.RollbackTrans
                Exit Function
            End If
        End If
        gstrSQL = "ZL_门诊_Clear('" & g病人身份_神木大兴.IC卡号 & "')"
        ExecuteProcedure_神木大兴 "清除数据"
    End If
    
    
    
    '第二步:上传明细数据
     DebugTool "门诊虚拟结算,第二步:上传明细数据"
    
    With rs明细
        If rs明细.RecordCount = 0 Then ShowMsgbox "未输入相关的费用记录!": Exit Function
        Do While Not .EOF
        
            '判断编码
            If Check医保项目(Nvl(!收费细目ID, 0), str类别, str医保编码, str医保名称, str拼音码) = False Then
                gcnOracle_神木大兴.RollbackTrans
                Exit Function
            End If
            '当前门诊号          As String       'right(病人ID1,6)-yyyyMMDDHHMMSS
            '明细参数:医保卡号_IN，流水号_IN，处方编号_IN，类别_IN，医保编码_IN，医保名称_IN，拼音编码_IN，价格IN，数量_IN，有效标志_IN
            
            gstrSQL = "ZL_MZ_FYMXB_INSERT("
            gstrSQL = gstrSQL & "'" & g病人身份_神木大兴.IC卡号 & "',"       '医保卡号_IN
            gstrSQL = gstrSQL & rs明细.AbsolutePosition & ","   '流水号_IN
            gstrSQL = gstrSQL & "'" & g病人身份_神木大兴.当前门诊号 & "',"    '处方编号_IN
            gstrSQL = gstrSQL & "'" & str类别 & "',"    '类别_IN
            gstrSQL = gstrSQL & "'" & str医保编码 & "',"   '医保编码_IN
            gstrSQL = gstrSQL & "'" & str医保名称 & "',"   '医保名称_IN
            If str拼音码 = "" Then
                gstrSQL = gstrSQL & "'" & zlCommFun.zlGetSymbol(str医保名称, 0) & "',"     '拼音编码_IN
            Else
                gstrSQL = gstrSQL & "'" & str拼音码 & "',"     '拼音编码_IN
            End If
            gstrSQL = gstrSQL & Format(Nvl(!实收金额, 0) / Nvl(!数量, 0), "####0.0000;-####0.0000;0;0") & "," '价格IN
            gstrSQL = gstrSQL & Format(Nvl(!数量, 0), "####0.00;-####0.00;0;0") & ","      '数量_IN
            gstrSQL = gstrSQL & "'F')"                                                              '有效标志_IN :医院MIS初值为F,只能由医保修改其标志，当医保正常接收完成后将其置为T,异常接收时将其标记值为X以便医院MIS查询?医院MIS无权修改其标志，否则后果自负
            DebugTool gstrSQL
            ExecuteProcedure_神木大兴 "门诊明细信息写入"
            g病人身份_神木大兴.费用总额 = g病人身份_神木大兴.费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
    
    '写结算信息:
    '  参数:  YBKH_IN IN MZ_JSLSB.YBKH%TYPE,--医保卡号     "
    '    CFBH_IN IN MZ_JSLSB.CFBH%TYPE,--处方编号
    '    FYHJ_IN IN MZ_JSLSB.FYHJ%TYPE,--费用合计
    '    XM_IN IN MZ_JSLSB.XM%TYPE   --姓名
    g病人身份_神木大兴.费用总额 = Round(g病人身份_神木大兴.费用总额, 2)
    
    gstrSQL = "ZL_MZ_JSLSB_INSERT("
    gstrSQL = gstrSQL & "'" & g病人身份_神木大兴.IC卡号 & "',"
    gstrSQL = gstrSQL & "'" & g病人身份_神木大兴.当前门诊号 & "')"
    
    ExecuteProcedure_神木大兴 "门诊结算信息写入"
   DebugTool "门诊虚拟结算,写结算信息"
    
    gcnOracle_神木大兴.CommitTrans
    
    
    '第三步:等待结算信息
     DebugTool "门诊虚拟结算,第三步:等待结算信息"
    
    If frm请求等待_神木大兴.ShowWait(0, g病人身份_神木大兴.IC卡号) = False Then
        gcnOracle_神木大兴.BeginTrans
        gstrSQL = "ZL_门诊_Clear('" & g病人身份_神木大兴.IC卡号 & "')"
        ExecuteProcedure_神木大兴 "清除数据"
        gcnOracle_神木大兴.CommitTrans
        Exit Function
    End If
        
    gstrSQL = "" & _
       "   Select  ybkh 医保卡号, cfbh 处方编号, jssj 结算时间, jsbz 医保结算标志, " & _
       "           fyhj 本次总费用, kszf 卡上支付, tczf 统筹支付, ybje 应补现金额, xm 病人姓名,jsqksye 结算前卡上余额,jshksye 结算后卡上余额 " & _
       "   From MZ_JSLSB  " & _
       "   Where ybkh='" & g病人身份_神木大兴.IC卡号 & "' and jsbz='T'"
    
    OpenRecordset_神木大兴 rsTemp, "获取结算信息", gstrSQL
    str结算方式 = ""
    With g结算数据
        .个人帐户支付 = Format(Nvl(rsTemp!卡上支付, 0), "####0.00;-####0.00;0;0")
        .统筹支付 = Format(Nvl(rsTemp!统筹支付, 0), "####0.00;-####0.00;0;0")
        .应补现金额 = Format(Nvl(rsTemp!应补现金额, 0), "####0.00;-####0.00;0;0")
        .费用总额 = Format(Nvl(rsTemp!本次总费用, 0), "####0.00;-####0.00;0;0")
        .结算前卡上余额 = Format(Nvl(rsTemp!结算前卡上余额, 0), "####0.00;-####0.00;0;0")
        .结算后卡上余额 = Format(Nvl(rsTemp!结算后卡上余额, 0), "####0.00;-####0.00;0;0")
        .公费床位费 = 0
        .公费调温费 = 0
        .公务员床补 = 0
        .押金总额 = 0
        .自费床位费 = 0
        .自费调温费 = 0
        str结算方式 = "个人帐户;" & .个人帐户支付 & ";0"
        str结算方式 = str结算方式 & "|" & "统筹支付;" & .统筹支付 & ";0"
    End With
    DebugTool "门诊虚拟结算成功,结算方式：" & str结算方式
    
    g病人身份_神木大兴.虚拟结算 = True
    门诊虚拟结算_神木大兴 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle_神木大兴.RollbackTrans
End Function
Public Function 门诊结算_神木大兴(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
        '此时所有收费细目必然有对应的医保编码
    Dim lng病人ID  As Long
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHandle

    Call DebugTool("进入门诊结算")

    gstrSQL = "" & _
        "   Select a.*,a.付数*a.数次 as 数量,a.实收金额/(nvl(a.付数,1)*nvl(a.数次,1)) as 单价 " & _
        "   From 门诊费用记录 a " & _
        "   Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"

    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取明细记录", lng结帐ID)

    If rs明细.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "没有填写收费记录!"
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")


    If g病人身份_神木大兴.病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人还没有经过身份验证，不能进行医保结算。"
        Exit Function
    End If
    g病人身份_神木大兴.结帐ID = lng结帐ID
    
    Dim dbl费用总额 As Double
    
    dbl费用总额 = 0
    '第一步:汇总费用
    DebugTool "门诊结算,第一步:汇总费用,及打上标志"
    With rs明细
        If rs明细.RecordCount = 0 Then ShowMsgbox "未输入相关的费用记录!": Exit Function
        Do While Not .EOF
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & g病人身份_神木大兴.当前门诊号 & "')"
            DebugTool "     打上明细标志:SQL=" & gstrSQL
            zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            DebugTool " 打上明细标志:更新病人费用记录成功:SQL=" & gstrSQL
            dbl费用总额 = dbl费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
    
    If Format(dbl费用总额, "#####0.00;-####0.00;0;0") <> Format(g结算数据.费用总额, "#####0.00;-####0.00;0;0") Then
        Err.Raise 9000, gstrSysName, "费用总额不等,不能结算!" & vbCrLf & _
                " 虚拟结算费用总额:" & Format(g结算数据.费用总额, "#####0.00;-####0.00; ;") & vbCrLf & _
                " 正式结算费用总额:" & Format(dbl费用总额, "#####0.00;-####0.00; ;")
        Exit Function
    End If
    
    
    '第二步:保存结算信息
    DebugTool "第二步:并开始保存保险结算记录"

   '插入保险结算记录
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(结算前卡上余额),累计统筹报销_IN(结算后卡上余额),住院次数_IN(住院:主页id),起付线(押金总额),封顶线_IN(公费床位费),实际起付线_IN(自费床位费),
    '   发生费用金额_IN(费用总金额),全自付金额_IN(公费调温费),首先自付金额_IN(应补现金额),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(公务员床补),超限自付金额_IN(自费调温费),个人帐户支付_IN(个人帐户支付金额),"
    '   支付顺序号_IN(门诊:处方号),主页ID_IN(主页id),中途结帐_IN,备注_IN()
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_陕西大兴 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "NULL,NULL," & g结算数据.结算前卡上余额 & "," & g结算数据.结算后卡上余额 & ",null,0,0,0," & _
            g结算数据.费用总额 & ",0," & g结算数据.应补现金额 & "," & _
            g结算数据.统筹支付 & " ," & g结算数据.统筹支付 & ",0,0," & g结算数据.个人帐户支付 & ",'" & _
             g病人身份_神木大兴.当前门诊号 & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
  
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_陕西大兴 & ",'帐户余额','" & g结算数据.结算后卡上余额 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存最后一次的卡上余额")
  
    '第三步:清除中间数据信息
    DebugTool "第三步:清除中间数据信息"
    
    gstrSQL = "ZL_门诊_Clear('" & g病人身份_神木大兴.IC卡号 & "')"
    ExecuteProcedure_神木大兴 "清除数据"
    
    门诊结算_神木大兴 = True
    g病人身份_神木大兴.虚拟结算 = False
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_神木大兴(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额

    Dim intMouse As Integer
    Dim lng冲销ID  As Long
    Dim str类别 As String, str医保编码 As String, str医保名称 As String, str拼音码 As String
    Dim rs明细 As New ADODB.Recordset
    Dim rs原明细 As New ADODB.Recordset

    Dim rsTemp As New ADODB.Recordset
    Dim lng病人id1 As Long
On Error GoTo errHand:

    门诊结算冲销_神木大兴 = False

    '身份验证
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If 身份标识_神木大兴(2, lng病人id1) = "" Then
        If lng病人id1 = 0 Then
            Err.Raise 9000, gstrSysName, "你不是当前持卡人!"
            Screen.MousePointer = intMouse
            Exit Function
        End If
    End If
    If lng病人ID <> lng病人id1 Then
        Screen.MousePointer = intMouse
        Err.Raise 9000, gstrSysName, "你不是当前持卡人!"
        Exit Function
    End If

    Err = 0:
    Screen.MousePointer = intMouse

    gcnOracle_神木大兴.BeginTrans


    '第一步:确定冲销ID值
    DebugTool "第一步:确定冲销ID值"
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")

    '第二步:确定冲销和原始单据的明细记录
    DebugTool "确定冲销和原始单据的明细记录"

    gstrSQL = "Select * From 门诊费用记录 " & _
        " Where 结帐ID=" & lng冲销ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0 and nvl(实收金额,0)<>0"

    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取冲销记录")
    g病人身份_神木大兴.费用总额 = 0
    g病人身份_神木大兴.结帐ID = lng结帐ID
     


    gstrSQL = "Select * From 门诊费用记录 where  结帐ID = [1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0 and nvl(实收金额,0)<>0"
    Set rs原明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取冲销记录", lng结帐ID)
    g病人身份_神木大兴.当前门诊号 = Nvl(rs原明细!摘要)



    '第三步:将原始记录中的接要作为冲销单据的接要，并打上上传标志
    DebugTool "第三步:将原始记录中的接要作为冲销单据的接要，并打上上传标志"
    With rs明细
        Do While Not .EOF
        
            '判断编码
            If Check医保项目(Nvl(!收费细目ID, 0), str类别, str医保编码, str医保名称, str拼音码) = False Then
                gcnOracle_神木大兴.RollbackTrans
                Exit Function
            End If
        
            '当前门诊号          As String       'right(病人ID1,6)-yyyyMMDDHHMMSS
            '明细参数:医保卡号_IN，流水号_IN，处方编号_IN，类别_IN，医保编码_IN，医保名称_IN，拼音编码_IN，价格IN，数量_IN，有效标志_IN
            
            gstrSQL = "ZL_MZ_FYMXB_INSERT("
            gstrSQL = gstrSQL & "'" & g病人身份_神木大兴.IC卡号 & "',"       '医保卡号_IN
            gstrSQL = gstrSQL & .AbsolutePosition & ","   '流水号_IN
            gstrSQL = gstrSQL & "'" & g病人身份_神木大兴.当前门诊号 & "',"    '处方编号_IN
            gstrSQL = gstrSQL & "'" & str类别 & "',"    '类别_IN
            gstrSQL = gstrSQL & "'" & str医保编码 & "',"   '医保编码_IN
            gstrSQL = gstrSQL & "'" & str医保名称 & "',"   '医保名称_IN
            If str拼音码 = "" Then
                gstrSQL = gstrSQL & "'" & zlCommFun.zlGetSymbol(str医保名称, 0) & "',"     '拼音编码_IN
            Else
                gstrSQL = gstrSQL & "'" & str拼音码 & "',"     '拼音编码_IN
            End If
            gstrSQL = gstrSQL & Format(Nvl(!实收金额, 0) / (Nvl(!付数, 1) * Nvl(!数次, 1)), "####0.0000;-####0.0000;0;0") & "," '价格IN
            gstrSQL = gstrSQL & Format((Nvl(!付数, 1) * Nvl(!数次, 1)), "####0.00;-####0.00;0;0") & ","      '数量_IN
            gstrSQL = gstrSQL & "'F')"                                                              '有效标志_IN :医院MIS初值为F,只能由医保修改其标志，当医保正常接收完成后将其置为T,异常接收时将其标记值为X以便医院MIS查询?医院MIS无权修改其标志，否则后果自负
            DebugTool gstrSQL
            ExecuteProcedure_神木大兴 "门诊明细信息写入"
            
            '写上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & g病人身份_神木大兴.当前门诊号 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
            g病人身份_神木大兴.费用总额 = g病人身份_神木大兴.费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With

    
    '第四步:冲销的相关记录
    DebugTool "第四步:冲销相关记录"
    
    gstrSQL = "Select * from 保险结算记录 where 性质=1 and 记录id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取中心单据号"


   '插入保险结算记录
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(),累计统筹报销_IN(),住院次数_IN(住院:主页id),起付线(押金总额),封顶线_IN(公费床位费),实际起付线_IN(自费床位费),
    '   发生费用金额_IN(费用总金额),全自付金额_IN(公费调温费),首先自付金额_IN(应补现金额),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(公务员床补),超限自付金额_IN(自费调温费),个人帐户支付_IN(个人帐户支付金额),"
    '   支付顺序号_IN(门诊:处方号),主页ID_IN(主页id),中途结帐_IN,备注_IN()
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_陕西大兴 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL,NULL,null,null," & -1 * Nvl(rsTemp!起付线, 0) & "," & -1 * Nvl(rsTemp!封顶线, 0) & "," & -1 * Nvl(rsTemp!实际起付线, 0) & "," & _
           -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & "," & -1 * Nvl(rsTemp!首先自付金额, 0) & "," & _
           -1 * Nvl(rsTemp!进入统筹金额, 0) & "," & -1 * Nvl(rsTemp!统筹报销金额, 0) & "," & -1 * Nvl(rsTemp!大病自付金额, 0) & "," & -1 * Nvl(rsTemp!超限自付金额, 0) & "," & -1 * Nvl(rsTemp!个人帐户支付, 0) & ",'" & _
           Nvl(rsTemp!支付顺序号, 0) & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    gcnOracle_神木大兴.CommitTrans
    
    门诊结算冲销_神木大兴 = True
    Exit Function
errHand:
    gcnOracle_神木大兴.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function


Public Function 入院登记_神木大兴(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    
    
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select to_char(入院日期,'yyyy-mm-dd hh24:mi:ss') as 入院日期 From 病案主页 where 病人id= " & lng病人ID & " and 主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取入院日期"
     
   gcnOracle_神木大兴.BeginTrans
    
    '过程参数
    '   医保卡号_IN,住院编号_IN,入院时间_IN,结算标志_IN
    gstrSQL = "ZL_ZY_JSLSB_INSERT("
    gstrSQL = gstrSQL & "'" & g病人身份_神木大兴.IC卡号 & "',"
    gstrSQL = gstrSQL & "'" & lng病人ID & "_" & lng主页ID & "',"
    gstrSQL = gstrSQL & "to_date('" & Nvl(rsTemp!入院日期) & "','yyyy-mm-dd hh24:mi:ss'),"
    gstrSQL = gstrSQL & "'F')"
    ExecuteProcedure_神木大兴 "更新病人入院"
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_陕西大兴 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
   gcnOracle_神木大兴.CommitTrans
    入院登记_神木大兴 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle_神木大兴.RollbackTrans
    入院登记_神木大兴 = False
End Function

Public Function 入院登记撤销_神木大兴(lng病人ID As Long, lng主页ID As Long) As Boolean
  '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
            
    '刘兴宏:20040923增加的
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
     Err = 0
    On Error GoTo errHand
    
    DebugTool "进入扩院登撤消接口"
    
    入院登记撤销_神木大兴 = False
    
    If 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "存在未结费用，不能撤消入院登记"
        Exit Function
    End If
    gstrSQL = "Select 医保号 From 保险帐户 where 病人id=" & lng病人ID & " and 险类=" & TYPE_陕西大兴
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取保险的相关信息"
    If rsTemp.EOF Then
        ShowMsgbox "不存在该医保病人!"
        Exit Function
    End If

    
    '参数为:医保卡号_IN
    gstrSQL = "ZL_ZY_JSLSB_DELETE("
    gstrSQL = gstrSQL & "'" & Nvl(rsTemp!医保号) & "')"
    ExecuteProcedure_神木大兴 "删除入院信息"
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_陕西大兴 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
        
    '更新医保帐户
    DebugTool "取消成功"
    入院登记撤销_神木大兴 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_神木大兴(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    Err = 0:    On Error GoTo errHand:
    出院登记_神木大兴 = False
    
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "当前病人不存在未结费用，请在入院撤消即可"
        Exit Function
    End If
    
    gstrSQL = "Select 医保号 From 保险帐户 where 病人id=" & lng病人ID & " and 险类=" & TYPE_陕西大兴
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取保险的相关信息"
    If rsTemp.EOF Then
        ShowMsgbox "不存在该医保病人!"
        Exit Function
    End If
    
    '参数:医保卡号_IN,结算标志_IN
    gstrSQL = "ZL_ZY_JSLSB_UPDATE("
    gstrSQL = gstrSQL & "'" & Nvl(rsTemp!医保号) & "',"
    gstrSQL = gstrSQL & "'T')"
    ExecuteProcedure_神木大兴 "更新病人出院标志"
    '改变当前状态
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_陕西大兴 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_神木大兴 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function 出院登记撤销_神木大兴(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
  '出院登记撤消
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    
    出院登记撤销_神木大兴 = False
    
    Err = 0: On Error GoTo errHand:
     
     If Not 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "该病人已经出院结算了,不能再取消出院!"
        Exit Function
     End If
    
    gstrSQL = "Select 医保号 From 保险帐户 where 病人id=" & lng病人ID & " and 险类=" & TYPE_陕西大兴
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取保险的相关信息"
    If rsTemp.EOF Then
        ShowMsgbox "不存在该医保病人!"
        Exit Function
    End If
    
    '参数:医保卡号_IN,结算标志_IN
    gstrSQL = "ZL_ZY_JSLSB_UPDATE("
    gstrSQL = gstrSQL & "'" & Nvl(rsTemp!医保号) & "',"
    gstrSQL = gstrSQL & "'F')"
    ExecuteProcedure_神木大兴 "更新病人出院标志"
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_陕西大兴 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    出院登记撤销_神木大兴 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function GetMax流水号(ByVal str卡号 As String) As Long
    Dim rsTemp As New ADODB.Recordset
    
    '获取最大流水号
    gstrSQL = "Select nvl(max(ID),0)+1  as 序号 From ZY_FYMXB where YBKH='" & str卡号 & "'"
    OpenRecordset_神木大兴 rsTemp, "获取最大号", gstrSQL
    If rsTemp.EOF Then
        GetMax流水号 = 1
    Else
       GetMax流水号 = Nvl(rsTemp!序号, 1)
    End If
    
    
End Function
Public Function 处方登记_神木大兴(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str流水号 As String
    Dim str类别 As String, str医保编码 As String, str医保名称 As String, str拼音码 As String


    处方登记_神木大兴 = False


   '读出该张单据的费用明细
   gstrSQL = "" & _
              "  Select A.*,M.医保号" & _
              "  From 住院费用记录 A,病案主页 C,保险帐户 M" & _
              "  where a.NO=[1] and A.记录性质=[2] and A.记录状态 = [3]" & _
              "        and A.病人ID=C.病人ID and nvl(a.实收金额,0)<>0 and A.主页ID=C.主页ID  And a.病人id=M.病人id and M.险类=[4] and  C.险类=[4]" & _
              "  Order by A.病人ID,A.NO,A.发生时间"
    
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "处方明细上传", str单据号, lng记录性质, lng记录状态, TYPE_陕西大兴)
    Err = 0:    On Error GoTo errHand:
    With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
        
            If Check医保项目(Nvl(!收费细目ID, 0), str类别, str医保编码, str医保名称, str拼音码) = False Then Exit Function
                        
            '明细参数:医保卡号_IN,流水号_IN,住院编号_IN,类别_IN,医保编码_IN,医保名称_IN,拼音编码_IN,价格_IN,数量_IN,有效标志_IN
            str流水号 = GetMax流水号(Nvl(!医保号))
            gstrSQL = "ZL_ZY_FYMXB_INSERT("
            gstrSQL = gstrSQL & "'" & Nvl(!医保号) & "',"       '医保卡号_IN
            gstrSQL = gstrSQL & str流水号 & ","   '流水号_IN
            gstrSQL = gstrSQL & "'" & Nvl(!病人ID, 0) & "_" & Nvl(!主页ID, 0) & "',"   '住院编号_IN
            gstrSQL = gstrSQL & "'" & str类别 & "',"    '类别_IN
            gstrSQL = gstrSQL & "'" & str医保编码 & "',"   '医保编码_IN
            gstrSQL = gstrSQL & "'" & str医保名称 & "',"   '医保名称_IN
            If str拼音码 = "" Then
                gstrSQL = gstrSQL & "'" & zlCommFun.zlGetSymbol(str医保名称, 0) & "',"     '拼音编码_IN
            Else
                gstrSQL = gstrSQL & "'" & str拼音码 & "',"     '拼音编码_IN
            End If
            gstrSQL = gstrSQL & Format(Nvl(!实收金额, 0) / (Nvl(!付数, 1) * Nvl(!数次, 1)), "####0.0000;-####0.0000;0;0") & "," '价格IN
            gstrSQL = gstrSQL & Format((Nvl(!付数, 1) * Nvl(!数次, 1)), "####0.00;-####0.00;0;0") & ","   '数量_IN
            gstrSQL = gstrSQL & "'F')"                                                              '有效标志_IN :医院MIS初值为F,只能由医保修改其标志，当医保正常接收完成后将其置为T,异常接收时将其标记值为X以便医院MIS查询?医院MIS无权修改其标志，否则后果自负
            DebugTool gstrSQL
            ExecuteProcedure_神木大兴 "住院明细信息写入"
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str流水号 & "')"
            zlDatabase.ExecuteProcedure gstrSQL, "更新明细数据"
            .MoveNext
        Loop
    End With
    处方登记_神木大兴 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 补传住院明细记录_神木大兴(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional bln忽略 As Boolean = False) As Boolean

    '补传相关明细记录
    Dim cnTemp As New ADODB.Connection
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str流水号 As String
    Dim str类别 As String, str医保编码 As String, str医保名称 As String, str拼音码 As String
    
    Err = 0:    On Error GoTo errHand:
      
      
    补传住院明细记录_神木大兴 = False
    
   gstrSQL = "" & _
              "  Select A.*,M.医保号" & _
              "  From 住院费用记录 A,病案主页 C,保险帐户 M" & _
              "  where Nvl(A.是否上传,0)=0 And Nvl(附加标志,0)<>9  and a.病人id=[1] and A.主页id= [2]" & _
              "        and  A.病人ID=C.病人ID and nvl(a.实收金额,0)<>0 and A.主页ID=C.主页ID  And a.病人id=M.病人id and M.险类=[3] and  C.险类=[3]" & _
              "  Order by A.病人ID,A.NO,A.发生时间"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID, lng主页ID, TYPE_陕西大兴)

   With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            
            If Check医保项目(Nvl(!收费细目ID, 0), str类别, str医保编码, str医保名称, str拼音码, bln忽略) = False Then
                If bln忽略 = False Then Exit Function
            Else
                        
                '明细参数:医保卡号_IN,流水号_IN,住院编号_IN,类别_IN,医保编码_IN,医保名称_IN,拼音编码_IN,价格_IN,数量_IN,有效标志_IN
                str流水号 = GetMax流水号(Nvl(!医保号))
                gstrSQL = "ZL_ZY_FYMXB_INSERT("
                gstrSQL = gstrSQL & "'" & Nvl(!医保号) & "',"       '医保卡号_IN
                gstrSQL = gstrSQL & str流水号 & ","   '流水号_IN
                gstrSQL = gstrSQL & "'" & lng病人ID & "_" & lng主页ID & "',"     '住院编号_IN
                gstrSQL = gstrSQL & "'" & str类别 & "',"    '类别_IN
                gstrSQL = gstrSQL & "'" & str医保编码 & "',"   '医保编码_IN
                gstrSQL = gstrSQL & "'" & str医保名称 & "',"   '医保名称_IN
                If str拼音码 = "" Then
                    gstrSQL = gstrSQL & "'" & zlCommFun.zlGetSymbol(str医保名称, 0) & "',"     '拼音编码_IN
                Else
                    gstrSQL = gstrSQL & "'" & str拼音码 & "',"     '拼音编码_IN
                End If
                gstrSQL = gstrSQL & Format(Nvl(!实收金额, 0) / (Nvl(!付数, 1) * Nvl(!数次, 1)), "####0.0000;-####0.0000;0;0") & "," '价格IN
                gstrSQL = gstrSQL & Format((Nvl(!付数, 1) * Nvl(!数次, 1)), "####0.00;-####0.00;0;0") & ","   '数量_IN
                gstrSQL = gstrSQL & "'F')"                                                              '有效标志_IN :医院MIS初值为F,只能由医保修改其标志，当医保正常接收完成后将其置为T,异常接收时将其标记值为X以便医院MIS查询?医院MIS无权修改其标志，否则后果自负
                DebugTool gstrSQL
                ExecuteProcedure_神木大兴 "住院明细信息写入"
                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str流水号 & "')"
                zlDatabase.ExecuteProcedure gstrSQL, "更新明细数据"
            End If
            .MoveNext
        Loop
    End With
    补传住院明细记录_神木大兴 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_神木大兴(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String
    'rsExse:字符集
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset

    Dim lng主页ID As Long, StrInput As String, strOutput  As String
    Dim str住院号 As String, str结算方式 As String, strSQL As String
    Dim lng病人id1 As Long
    Dim intMouse As Integer

    Dim strArr As Variant

    Err = 0: On Error GoTo errHand:

    g病人身份_神木大兴.病人ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If
    'intMouse = Screen.MousePointer
    gstrSQL = "Select a.*,b.* From 保险帐户 a,病人信息 b where a.病人id=b.病人id and a.病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID)
    If rsTemp.EOF Then
        ShowMsgbox "无相关的医保证号!"
        Exit Function
    End If
    g病人身份_神木大兴.IC卡号 = Nvl(rsTemp!卡号)
    g病人身份_神木大兴.姓名 = Nvl(rsTemp!姓名)
    g病人身份_神木大兴.性别 = Nvl(rsTemp!性别)
    

    gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID)

    If IsNull(rsTemp("主页ID")) = True Then
        MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    lng主页ID = rsTemp("主页ID")


'    If bln结帐处 Then
'        Screen.MousePointer = 1
'        If 身份标识_神木大兴(4, lng病人id1) = "" Then
'            Screen.MousePointer = intMouse
'            住院虚拟结算_神木大兴 = ""
'            Exit Function
'        End If
'        Screen.MousePointer = intMouse
'        If lng病人ID <> lng病人id1 Then
'            ShowMsgbox "不是当前要结算的病人!"
'            Exit Function
'        End If
'    End If

    
    'Screen.MousePointer = vbHourglass

    
    g病人身份_神木大兴.结帐ID = 0
    g病人身份_神木大兴.病人ID = lng病人ID
    
    
    '第一步:汇总费用
    DebugTool "住院虚拟结算,第一步:汇总费用"
    g病人身份_神木大兴.费用总额 = 0
    With rsExse
        If rsExse.RecordCount = 0 Then ShowMsgbox "未输入相关的费用记录!": Exit Function
        Do While Not .EOF
            g病人身份_神木大兴.费用总额 = g病人身份_神木大兴.费用总额 + Nvl(!金额, 0)
            .MoveNext
        Loop
    End With
    
    '第二步:补传明细费用
    DebugTool "住院虚拟结算,第二步:补传明细费用"
    If 补传住院明细记录_神木大兴(lng病人ID, lng主页ID) = False Then Exit Function
    
    '第三步:更新总费用
    
    '第三步:等待结算信息
     DebugTool "住院虚拟结算,第三步:等待结算信息"
    
    If frm请求等待_神木大兴.ShowWait(1, g病人身份_神木大兴.IC卡号) = False Then
        Exit Function
    End If

    
    
    '第四步:分解相关结果
     DebugTool "住院虚拟结算,第四步:分解相关结果"

     gstrSQL = "" & _
        "   Select  ybkh 医保卡号, zybh 住院编号, rysj 入院时间, cysj 结算时间, jsbz 医保结算标志, tpbz 医保退票标志, " & _
        "           yybz 医院结算标志, fyhj 本次总费用, kszf 卡上支付, tczf 统筹支付, gwycb 公务员床补," & _
        "           yj 押金总额, ybje 应补现金额, gfcwf 公费床位费, zfcwf 自费床位费, gftwf 公费调温费, zftwf 自费调温费 " & _
        "   from zy_jslsb   " & _
        "   where ybkh='" & g病人身份_神木大兴.IC卡号 & "'"
     Call OpenRecordset_神木大兴(rsTemp, "获取住院结算信息", gstrSQL)
    str结算方式 = ""
    With g结算数据
        .个人帐户支付 = Format(Nvl(rsTemp!卡上支付, 0), "####0.00;-####0.00;0;0")
        .统筹支付 = Format(Nvl(rsTemp!统筹支付, 0), "####0.00;-####0.00;0;0")
        .应补现金额 = Format(Nvl(rsTemp!应补现金额, 0), "####0.00;-####0.00;0;0")
        .公费床位费 = Format(Nvl(rsTemp!公费床位费, 0), "####0.00;-####0.00;0;0")
        .公费调温费 = Format(Nvl(rsTemp!公费调温费, 0), "####0.00;-####0.00;0;0")
        .公务员床补 = Format(Nvl(rsTemp!公务员床补, 0), "####0.00;-####0.00;0;0")
        .押金总额 = Format(Nvl(rsTemp!押金总额, 0), "####0.00;-####0.00;0;0")
        .自费床位费 = Format(Nvl(rsTemp!自费床位费, 0), "####0.00;-####0.00;0;0")
        .自费调温费 = Format(Nvl(rsTemp!自费调温费, 0), "####0.00;-####0.00;0;0")
        .费用总额 = Format(Nvl(rsTemp!本次总费用, 0), "####0.00;-####0.00;0;0")
        
        
        str结算方式 = "个人帐户;" & .个人帐户支付 & ";0"
        str结算方式 = str结算方式 & "|" & "统筹支付;" & .统筹支付 & ";0"
        str结算方式 = str结算方式 & "|" & "公务员床补;" & .公务员床补 & ";0"
    End With
    DebugTool "住院虚拟结算成功,结算方式：" & str结算方式


    住院虚拟结算_神木大兴 = str结算方式
    g病人身份_神木大兴.病人ID = lng病人ID   '表示该病人已经进行了虚拟结算
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function




Public Function 住院结算_神木大兴(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)

    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput As String

    Dim lng主页ID As Long
    Dim dbl费用总额 As Double
    Dim strArr As Variant, strTmpArr As Variant

    Dim str结算方式  As String, str住院号 As String
    Dim obj结算 As 结算数据
    Dim dbl个人帐户 As Double

    住院结算_神木大兴 = False


    Err = 0: On Error GoTo errHand:
    Call DebugTool("进入住院结算")

    
    '第一步:检查数据的正误
    DebugTool "第一步:检查数据的正误"
    If g病人身份_神木大兴.病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人没有完成医保的预结算操作，不能进行结算。"
        Exit Function
    End If

    gstrSQL = "Select 当前状态 From 保险帐户  where 病人ID=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "判断当前的住院状态!"

    If Nvl(rsTemp!当前状态, 0) = 1 Then
        Err.Raise 9000, gstrSysName, "当前病人还处于在院状态,请出院后再结算!"
        Exit Function
    End If


    With g结算数据
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID)
        If IsNull(rsTemp("主页ID")) = True Then
            Err.Raise 9000, gstrSysName, "只有住院病人才可以使用医保结算。"
            Exit Function
        End If
        lng主页ID = rsTemp("主页ID")
    End With

    gstrSQL = " " & _
          " Select sum(round(nvl(结帐金额,0),2)) as 实收金额 " & _
          " From 住院费用记录 " & _
          " Where 记录状态<>0 and 结帐ID=" & lng结帐ID & " and  Nvl(附加标志,0)<>9"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取总费用"

    dbl费用总额 = Round(Val(Nvl(rsTemp!实收金额, 0)), 2)
    If dbl费用总额 <> Round(g结算数据.费用总额, 2) Then
        If dbl费用总额 - Round(g结算数据.费用总额, 2) <= 0.1 Then
            If MsgBox("虚拟结算数据的费用总额(" & Format(g结算数据.费用总额, "####0.00;-###0.00;0;0") & ")" & vbCrLf & "与本次结算的费用总额(" & Format(dbl费用总额, "####0.00;-###0.00;0;0") & ")不等，是否继续结算?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            Err.Raise 9000, gstrSysName, "虚拟结算数据的费用总额(" & Format(g结算数据.费用总额, "####0.00;-###0.00;0;0") & ")" & vbCrLf & "与本次结算的费用总额(" & Format(dbl费用总额, "####0.00;-###0.00;0;0") & ")不等，请检查处方是否正确!"
            Exit Function
        End If
    End If
    g病人身份_神木大兴.结帐ID = lng结帐ID
    
    
    '第二步:保存结算信息
    DebugTool "第二步:并开始保存保险结算记录"

   '插入保险结算记录
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(),累计统筹报销_IN(),住院次数_IN(住院:主页id),起付线(押金总额),封顶线_IN(公费床位费),实际起付线_IN(自费床位费),
    '   发生费用金额_IN(费用总金额),全自付金额_IN(公费调温费),首先自付金额_IN(应补现金额),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(公务员床补),超限自付金额_IN(自费调温费),个人帐户支付_IN(个人帐户支付金额),"
    '   支付顺序号_IN(门诊:处方号),主页ID_IN(主页id),中途结帐_IN,备注_IN()
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_陕西大兴 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "NULL,NULL,NULL,NULL," & lng主页ID & "," & g结算数据.押金总额 & "," & g结算数据.公费床位费 & "," & g结算数据.自费床位费 & "," & _
            g结算数据.费用总额 & "," & g结算数据.公费调温费 & "," & g结算数据.应补现金额 & "," & _
            g结算数据.统筹支付 & " ," & g结算数据.统筹支付 & "," & g结算数据.公务员床补 & "," & g结算数据.自费调温费 & "," & g结算数据.个人帐户支付 & ",'" & _
             lng病人ID & "_" & lng主页ID & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
  
    '第三步:清除中间数据信息
    DebugTool "第三步:清除中间数据信息"
    gstrSQL = "ZL_住院_Clear('" & g病人身份_神木大兴.IC卡号 & "')"
    ExecuteProcedure_神木大兴 "清除数据"
    住院结算_神木大兴 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Public Function 住院结算冲销_神木大兴(lng结帐ID As Long) As Boolean
     '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    Err.Raise 9000, gstrSysName, "本医保不支持住院结算冲销,具体请咨询接口商!"
    住院结算冲销_神木大兴 = False
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 医保设置_神木大兴(ByVal lng险类 As Long, ByVal lng医保中心 As Integer) As Boolean
    '功能： 该方法用于供相关应用部件调用配置连接医保数据服务器的连接串
    '返回：接口配置成功，返回true；否则，返回false
    
    Dim strConn As String
    Dim blnReturn As Boolean
    
    If frmSet神木大兴.参数设置 = False Then
        Exit Function
    End If
  
    If gcnOracle_神木大兴 Is Nothing Then
                blnReturn = True
    Else
        If Open中间库_神木大兴() Then
                blnReturn = True
        End If
    End If
    医保设置_神木大兴 = blnReturn
End Function
Public Sub ExecuteProcedure_神木大兴(ByVal strCaption As String)
    '功能：执行SQL语句
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_神木大兴.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub
Public Function Open中间库_神木大兴() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    Dim strConn As String
    
    Open中间库_神木大兴 = False
    Err = 0: On Error Resume Next
        
    Err = 0: On Error GoTo errHand:
    
    '重新建立到医保服务器的公共连接
    '中间库连接
    gstrSQL = "select 参数名,参数值 from 保险参数 where  险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "神木大兴核工业医保", TYPE_陕西大兴)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    
    Set gcnOracle_神木大兴 = New ADODB.Connection
    If OraDataOpen(gcnOracle_神木大兴, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    Open中间库_神木大兴 = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub OpenRecordset_神木大兴(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_神木大兴, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

