Attribute VB_Name = "mdlCardSquare"
Option Explicit
Public Enum g小数类型
    g_数量 = 0
    g_成本价
    g_售价
    g_金额
    g_折扣率
End Enum
Private Type m_小数位
    数量小数 As Integer
    成本价小数 As Integer
    零售价小数 As Integer
    金额小数 As Integer
    折扣率 As Integer
End Type
Public g_小数位数 As m_小数位
Public gBytMoney As Byte '收费分币处理方法

'小数格式化串
Public Type g_FmtString
    FM_数量 As String
    FM_成本价 As String
    FM_零售价 As String
    FM_金额 As String
    FM_折扣率 As String
End Type
Public Enum gCardEditType   '卡编辑类型
    gEd_发卡 = 0
    gEd_修改 = 1
    gEd_换卡 = 2
    gEd_补卡 = 3
    gEd_查询 = 4
    gEd_充值 = 5
    gEd_充值回退 = 6
    gEd_回收 = 7
    gEd_取消回收 = 8
    gEd_退卡 = 9
    gEd_取消退卡 = 10
End Enum
Public Type zlTyCustumRecordset
    rs收费类别 As ADODB.Recordset
    rs消费卡接口 As ADODB.Recordset
    rs收费类别汇总 As ADODB.Recordset
    rs分单类别汇总 As ADODB.Recordset
    dbl费用总额 As Double
    dblHIS最大消费额 As Double
    dbl已刷累计额 As Double
End Type
Public gblnShowCard As Boolean  '就诊卡号显示(true,显示卡号,false,加密显示)
Public gObjXFCards As clsCards  '专门针对消费卡的(要管理卡号)
Public gobjSquare As SquareCard
Public gobjPublicExpense As Object  '费用公共部件
Public gintPriceGradeStartType As Integer
Public gstr药品价格等级 As String
Public gstr卫材价格等级 As String
Public gstr普通价格等级 As String

Public grsStatic As zlTyCustumRecordset
Public gVbFmtString As g_FmtString
Public gOraFmtString As g_FmtString
Public gbln自动读取 As Boolean '当前是否为射频卡
Public gblnCardNoSHowPW As Boolean  '卡号显示密文
Public gDebug As Boolean '调试开关
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjDataBase As Object
Public gobjControl As Object
Public gstrLike As String  '项目匹配方法,%或空
Private Type Ty_TestDebug
    blndebug As Boolean
    objSquareCard As clsCard
    bytType  As Byte  '1-随机产生卡号,2-读取卡号
    strStartNo As String    '开始卡号
    bln补调交易 As Boolean
End Type
Public gTy_TestBug As Ty_TestDebug
Public gobjStartCards As Collection  '启动的刷卡对象集
Public gbln消费卡退费验卡 As Boolean
 
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"
Public gintFeePrecision As Integer    '费用小数精度
Public gstrFeePrecisionFmt As String '费用小数格式:0.00000
Public gblnOK As Boolean
'LED语音报价控制
Public gblnLED As Boolean '是否使用Led显示
Public gblnLedWelcome As Boolean '是否显示欢迎信息

'门诊收据
Public gbln收费发票 As Boolean '发卡是否用收费发票
Public gblnBill发卡 As Boolean '是否严格票据管理
Public glngShareUseID As Long  '收费共享批次
Public gbyt收费 As Byte '收费票据长度
Public gblnStartFactUseType As Boolean '是否区分了使用类别
Public glngMax家庭地址 As Long       '家庭地址最大允许录入长度
Public glngMax户口地址 As Long       '户口地址最大允许录入长度
Public glngMax出生地点 As Long       '出生地点最大允许录入长度
Public glngMax联系人地址 As Long    '联系人地址最大允许录入长度
'Public gclsInsure As New clsInsure          '医保接口对象
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
    
    support挂号不收取病历费 = 81
End Enum

Public Sub zlinitSystemPara(Optional cnOracle As ADODB.Connection)
    '------------------------------------------------------------------------------
    '功能:初始化相关的系统参数
    '入参:cnOracle-数据库连接
    '返回:填充成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim strTemp As String, strValue As String
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objDatabase As Object, objTemp As clsDataBase
    
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDatabase = objTemp
    Else
         Set objDatabase = zlDatabase
    End If
    '问题:52913
    strSQL = "Select 卡号密文 From 医疗卡类别 Where 名称='就诊卡' and nvl(是否固定,0)=1"
    Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "读取原就诊卡卡号密文显示规则")
    
    gblnShowCard = False
    If Not rsTemp.EOF Then
        gblnShowCard = Nvl(rsTemp!卡号密文) = ""
    End If
    '104726:李南春,2017/4/24,收费发票打印发卡票据
    gbln收费发票 = Val(zlDatabase.GetPara("卡费使用门诊收费医疗收据", glngSys, glngModul)) = 1
    
    '票号严格控制
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnBill发卡 = Mid(strValue, 1, 1) = "1"
    '票据号码长度、就诊卡号长度
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    gbyt收费 = Val(Split(strValue, "|")(0))
    
    gbln消费卡退费验卡 = zlDatabase.GetPara(282, glngSys) = "1"
    
    '本地共用挂号批次ID
    If gbln收费发票 Then
        glngShareUseID = Val(zlDatabase.GetPara("共用门诊收据批次", glngSys, glngModul, ""))
        If glngShareUseID > 0 Then
            If Not ExistBill(glngShareUseID, 1) Then
                zlDatabase.SetPara "共用门诊收据批次", "0", glngSys, glngModul
                glngShareUseID = 0
            End If
        End If
    Else
        glngShareUseID = 0
    End If
    If gbln收费发票 Then
        gblnStartFactUseType = zlStartFactUseType("1")
    Else
        gblnStartFactUseType = False
    End If
    
    '78773:李南春,2014-10-29,LED显示一卡通支付信息
    gblnLED = Val(GetSetting("ZLSOFT", "公共全局", "使用", 0)) <> 0
    gblnLedWelcome = Val(objDatabase.GetPara("LED显示欢迎信息", glngSys, glngModul, 1)) <> 0
    gstrLike = IIf(Val(objDatabase.GetPara("输入匹配")) = 0, "%", "")
    With gSystemPara
        '0-拼音码,1-五笔码,2-两者
        .int简码方式 = Val(objDatabase.GetPara("简码方式"))
        .bln个性化风格 = objDatabase.GetPara("使用个性化风格") = "1"
        
        '第1位1-全数字只查编码,第2位1-全字母只查简码,在HIS基础参数中设置
        strTemp = objDatabase.GetPara(44, glngSys)
        If strTemp = "" Then strTemp = "00"
        If Len(strTemp) = 1 Then strTemp = strTemp & "0"
        .bln全数字按编码查 = Val(Left(strTemp, 1)) = 1
        .bln全字母按简码查 = Val(Mid(strTemp, 2, 1)) = 1
        '费用金额小数点位数
        gbytDec = Val(objDatabase.GetPara(9, glngSys, , 2))
        gstrDec = "0." & String(gbytDec, "0")
        '刘兴洪 问题:????    日期:2010-12-06 23:38:53
        '费用单价保留位数
        gintFeePrecision = Val(objDatabase.GetPara(157, glngSys, , "5"))
        gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
        '收费分币处理方式
        strValue = zlDatabase.GetPara(14, glngSys, , 0)
        gBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 2, 1)))
        
         .bln免挂号模式 = Val(zlDatabase.GetPara("免挂号模式", glngSys)) = 1
    
     End With
     gintDebug = -1
     '初如化站点信息
     Call Init站点信息: Call 初始小数位数
     Call zlInitColorSet: Call InitAddressLength
     Set objDatabase = Nothing
     Set objTemp = Nothing
End Sub
Public Sub 初始小数位数()
    '------------------------------------------------------------------------------------------------------
    '功能:初始小数位数
    '入参:
    '出参:
    '返回:7
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    With g_小数位数
        .成本价小数 = 7
        .零售价小数 = 7
        .金额小数 = 2
        .数量小数 = 3
        .折扣率 = 2
    End With
    With gVbFmtString
        .FM_成本价 = GetFmtString(g_成本价, False)
        .FM_金额 = GetFmtString(g_金额, False)
        .FM_零售价 = GetFmtString(g_售价, False)
        .FM_数量 = GetFmtString(g_数量, False)
        .FM_折扣率 = GetFmtString(g_折扣率, False)
    End With
    With gOraFmtString
        .FM_成本价 = GetFmtString(g_成本价, True)
        .FM_金额 = GetFmtString(g_金额, True)
        .FM_零售价 = GetFmtString(g_售价, True)
        .FM_数量 = GetFmtString(g_数量, True)
        .FM_折扣率 = GetFmtString(g_折扣率, True)
    End With
End Sub

Public Function GetFmtString(ByVal 小数类型 As g小数类型, Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '功能:返回指定的小数格式串
    '入参: lng小数位数-小数位数
    '     blnOracle-返回是oracle的格式串还是Vb的格式串
    '出参:
    '返回:返回指定的格式串
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strFmt As String
    Dim int位数 As Integer
    Select Case 小数类型
    Case g_数量
         int位数 = g_小数位数.数量小数
    Case g_金额
         int位数 = g_小数位数.金额小数
    Case g_成本价
         int位数 = g_小数位数.成本价小数
    Case g_售价
         int位数 = g_小数位数.零售价小数
    Case Else
        int位数 = 0
    End Select
    If blnOracle Then
       GetFmtString = "'999999999990." & String(int位数, "9") & "'"
    Else
       GetFmtString = "#0." & String(int位数, "0") & ";-#0." & String(int位数, "0") & "; ;"
    End If
End Function

Public Function zlGet收费类别() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费类别
    '编制:刘兴洪
    '日期:2009-12-09 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '先缓存到本地
    
    On Error GoTo errHandle
    
    gstrSQL = "Select  编码,名称 From 收费项目类别"
    If grsStatic.rs收费类别 Is Nothing Then
        Set grsStatic.rs收费类别 = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费类别")
    ElseIf grsStatic.rs收费类别.State <> 1 Then
        Set grsStatic.rs收费类别 = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费类别")
    End If
    grsStatic.rs收费类别.Filter = ""
    Set zlGet收费类别 = grsStatic.rs收费类别
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet消费卡接口(Optional cnOracle As ADODB.Connection, Optional ByVal blnOnlyStart As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡接口
    '入参:blnOnlyStart-是否仅读取启用的消费卡
    '编制:刘兴洪
    '日期:2009-12-09 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '先缓存到本地
    Dim objDatabase  As Object, objTemp As clsDataBase
    On Error GoTo errHandle
    Set objDatabase = zlDatabase
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDatabase = objTemp
    End If
    '56615
    '83399:李南春,2015/7/19,消费卡类别设置
    gstrSQL = _
        " Select 编号, 名称, 结算方式, Nvl(自制卡, 0) As 自制卡, 前缀文本, 卡号长度, 启用," & vbNewLine & _
        "        Nvl(是否退现, 0) As 是否退现, Nvl(是否全退, 0) As 是否全退," & vbNewLine & _
        "        Nvl(密码长度, 10) As 密码长度, Nvl(密码长度限制, 0) As 密码长度限制, Nvl(密码规则, 0) As 密码规则," & vbNewLine & _
        "        部件, 系统, 是否密文, 0 As 密码输入限制, 0 As 是否缺省密码," & vbNewLine & _
        "        0 As 是否模糊查找, 0 As 是否制卡, 1 As 是否发卡, 0 As 是否写卡, Nvl(读卡性质, '1000') As 读卡性质, nvl(键盘控制方式,0) As 键盘控制方式," & vbNewLine & _
        "        Nvl(是否严格控制, 0) As 是否严格控制, 限制类别, 应用场合, Nvl(是否特定病人, 0) As 是否特定病人," & vbNewLine & _
        "        Nvl(是否允许换卡, 0) As 是否允许换卡, Nvl(是否允许补卡, 0) As 是否允许补卡," & vbNewLine & _
        "        Nvl(是否允许余额退款, 0) As 是否允许余额退款" & vbNewLine & _
        " From 消费卡类别目录" & vbNewLine & _
        IIf(blnOnlyStart, " Where Nvl(启用, 0) = 1", "") & _
        " Order By 编号"
    If grsStatic.rs消费卡接口 Is Nothing Then
        Set grsStatic.rs消费卡接口 = objDatabase.OpenSQLRecord(gstrSQL, "获取消费卡接口 ")
    ElseIf grsStatic.rs消费卡接口.State <> 1 Then
        Set grsStatic.rs消费卡接口 = objDatabase.OpenSQLRecord(gstrSQL, "获取消费卡接口 ")
    End If

    grsStatic.rs消费卡接口.Filter = 0
    Set zlGet消费卡接口 = grsStatic.rs消费卡接口
    Exit Function
errHandle:
    If Not cnOracle Is Nothing And Not objTemp Is Nothing Then
        If objTemp.ErrCenter = 1 Then Resume
        Set objTemp = Nothing: Set objDatabase = Nothing
        Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Set objTemp = Nothing: Set objDatabase = Nothing
End Function

Public Function zlIsCardNoShowPW(ByRef lng序号 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示卡号是否密文显示
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-10-25 10:31:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = zlGet消费卡接口
    If rsTemp.EOF Then Exit Function
    rsTemp.Filter = "编号=" & lng序号
    If rsTemp.EOF Then
        zlIsCardNoShowPW = False
    Else
         zlIsCardNoShowPW = Val(Nvl(rsTemp!是否密文)) = 1
    End If
    rsTemp.Filter = 0
End Function
Public Function zlCreateBrushObjects(ByVal objCard As clsCard, ByRef objBrhushCardObject As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建刷卡对象
    '入参:clsCard-卡对象
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-31 14:46:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCommpentName As String
    If objCard.启用 Then
        '检查设备是否启用
        If objCard.接口程序名 = "" Then
            '消费卡
            Set objBrhushCardObject = New clsSimulateSquareCard: zlCreateBrushObjects = True
        Else
            strCommpentName = objCard.接口程序名 & "." & "cls" & Replace(Replace(UCase(objCard.接口程序名), "ZL9", ""), "ZL", "")
            Err = 0: On Error Resume Next
            Set objBrhushCardObject = CreateObject(strCommpentName)
            If Err <> 0 Then
                ShowMsgbox "部件:" & objCard.接口编码 & "-" & objCard.名称 & "( " & strCommpentName & ")创建失败!" & vbCrLf & "详细的信息为:" & Err.Description
                Call WritLog("mdlCardSquare.zlCreateBrushObjects", "", "部件:" & objCard.接口编码 & "-" & objCard.名称 & "创建失败!详细的信息为:" & Err.Description)
                Exit Function
            End If
            zlCreateBrushObjects = True
        End If
    Else
        Set objBrhushCardObject = Nothing
    End If
End Function
Public Function zlGetCardObject(ByVal lng接口编号 As Long, ByRef objBrushCard As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：根据指定结算卡序号获取结算卡对象
    '入参：lng接口编号-结算卡对序号
    '出参：objCard-返回结算卡对象
    '返回：获取成功,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-18 11:58:54
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objCardTemp As Object
    If gobjStartCards Is Nothing Then Exit Function
    
    If gobjStartCards.count = 0 Then Exit Function
    For i = 1 To gobjStartCards.count
         Err = 0: On Error Resume Next
         Set objCardTemp = gobjStartCards(i)(0)
         If Err = 0 Then
            If gobjStartCards(i)(2) = lng接口编号 Then
                Set objBrushCard = objCardTemp
                zlGetCardObject = True: Exit Function
            End If
        End If
        On Error GoTo 0
    Next
    zlGetCardObject = False
End Function

Public Function zlInitCards() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡集
    '返回:成功!返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-15 14:31:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, int自动读取 As Integer, bln启用 As Boolean, str部件 As String, objCard As clsCard
    Dim objBrushCards As Object, int自动间隔 As Integer
    
    Err = 0: On Error GoTo Errhand:
    Set gObjXFCards = New clsCards
    Set gobjStartCards = New Collection '格式;array(部件对象,自制卡,接口编号)
    Set rsTemp = zlGet消费卡接口
    With rsTemp
        '自制卡(即消费卡)
        .Filter = "自制卡=1"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
            int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & Nvl(!编号), "自动读取", "0"))
            bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & Nvl(!编号), "启用", "1")) = 1
            int自动间隔 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & Nvl(!编号), "自动读取间隔", "1"))
                
            str部件 = Nvl(rsTemp!部件)
            Set objCard = gObjXFCards.AddItem(EM_CardType_Consume, Val(Nvl(!编号)), Nvl(!编号), Nvl(rsTemp!名称), Left(Nvl(rsTemp!名称), 1), bln启用, True, str部件, True, 1, int自动读取, int自动间隔, Val(Nvl(rsTemp!系统)) = 1, Nvl(rsTemp!结算方式), Nvl(rsTemp!前缀文本), Val(Nvl(rsTemp!卡号长度)), True, Mid(Nvl(rsTemp!读卡性质), 1, 1) = 1, False, Val(Nvl(rsTemp!是否全退)) = 1, "", "", True, Val(Nvl(rsTemp!是否密文)), Val(Nvl(rsTemp!是否退现)) = 1, Val(Nvl(rsTemp!密码长度)), Val(Nvl(rsTemp!密码长度限制)), Val(Nvl(rsTemp!密码规则)), "K" & Nvl(rsTemp!编号))
            If zlCreateBrushObjects(objCard, objBrushCards) Then
                gobjStartCards.Add Array(objBrushCards, "1", CStr(Nvl(!编号))), "K" & Nvl(!编号)
            End If
            .MoveNext
        Loop
        '银联卡
        .Filter = "自制卡<>1"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & Nvl(!编号), "自动读取", 0))
            int自动间隔 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & Nvl(!编号), "自动读取间隔", "1"))
            bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & Nvl(!编号), "启用", "1")) = 1
            str部件 = Nvl(rsTemp!部件)
             Set objCard = gObjXFCards.AddItem(EM_CardType_Consume, Val(Nvl(!编号)), Nvl(!编号), Nvl(rsTemp!名称), Left(Nvl(rsTemp!名称), 1), bln启用, True, str部件, False, 1, int自动读取, int自动间隔, Val(Nvl(rsTemp!系统)) = 1, Nvl(rsTemp!结算方式), Nvl(rsTemp!前缀文本), Val(Nvl(rsTemp!卡号长度)), True, Mid(Nvl(rsTemp!读卡性质), 1, 1) = 1, True, Val(Nvl(rsTemp!是否全退)) = 1, "", "", True, Val(Nvl(rsTemp!是否密文)), Val(Nvl(rsTemp!是否退现)) = 1, Val(Nvl(rsTemp!密码长度)), Val(Nvl(rsTemp!密码长度限制)), Val(Nvl(rsTemp!密码规则)), "K" & Nvl(rsTemp!编号))
            If zlCreateBrushObjects(objCard, objBrushCards) Then
                gobjStartCards.Add Array(objBrushCards, 0, CStr(Nvl(!编号))), "K" & Nvl(!编号)
            End If
            .MoveNext
        Loop
    End With
    zlInitCards = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub WritLog(ByVal strDev As String, strInput As String, strOutPut As String)
    Call LogWrite("一卡通接口调试日志", glngModul, "读卡接口返回", "函数名:" & strDev & ";输入:" & strInput & ";输出:" & strOutPut)
End Sub

Public Function Read模拟卡号(ByVal strFile As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从已经产生的卡号中读取一个带标志的卡号(如果有多个,以最后一个为准)
    '编制:刘兴洪
    '日期:2009-12-17 10:35:51
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim objFile As New FileSystemObject, objText As TextStream, varData As Variant
    Dim strText As String, strCardNo As String
    strCardNo = ""
    Set objText = objFile.OpenTextFile(strFile)
    Do While Not objText.AtEndOfStream
        strText = Trim(objText.ReadLine)
        varData = Split(strText, vbTab)
        If Val(varData(0)) = 1 Then
            strCardNo = varData(1)
        End If
    Loop
    objText.Close
    Read模拟卡号 = strCardNo
    Exit Function
Errhand:
End Function
Public Sub zlInitBrushCardRec(ByRef rsTemp As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化本地记录集
    '出参:返回本地结算的初化记录休
    '编制:刘兴洪
    '日期:2009-12-23 11:22:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        If .State = adStateOpen Then .Close
        .Fields.Append "接口编号", adDouble, 18, adFldIsNullable
        .Fields.Append "消费卡ID", adDouble, 18, adFldIsNullable
        .Fields.Append "卡号", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "结算方式", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "卡名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "余额", adDouble, 16, adFldIsNullable
        .Fields.Append "结算金额", adDouble, 16, adFldIsNullable
        .Fields.Append "交易时间", adDate, 50, adFldIsNullable
        .Fields.Append "交易流水号", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "备注", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "结算标志", adNumeric, 2, adFldIsNullable
        .Fields.Append "分摊页码", adLongVarChar, 600, adFldIsNullable  '多单据有效,在HIS结算后自动分配:用逗号分离,如,2,3表示,此条刷卡消费分配在第二张单据和第三张单据
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
End Sub
Public Sub zlInit收费类别Struc()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化本地记录集
    '出参:返回本地结算的初化记录休
    '编制:刘兴洪
    '日期:2009-12-23 11:22:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set grsStatic.rs收费类别汇总 = New ADODB.Recordset
    Set grsStatic.rs分单类别汇总 = New ADODB.Recordset
    
    grsStatic.dbl费用总额 = 0: grsStatic.dbl已刷累计额 = 0
    With grsStatic.rs收费类别汇总
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "实收金额", adDouble, 16, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    
    With grsStatic.rs分单类别汇总
        If .State = adStateOpen Then .Close
        .Fields.Append "分类", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "单据序号", adDouble, 18, adFldIsNullable
        .Fields.Append "收费类别", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "实收金额", adDouble, 16, adFldIsNullable
        .Fields.Append "分摊金额", adDouble, 16, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
End Sub
Public Function zlInit收费类别数据(ByVal rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用记录集，获取当前卡可以消费的最大额度
    '入参:rsFeeList-明细费用:
    '    字段: 费别,NO,实际票号、结算时间、病人ID、收费类别、收据费目、计算单位、开单人、收费细目ID、数量、单价、实收金额、是否急诊、开单部门ID、执行部门ID
    '出参:
    '编制:刘兴洪
    '日期:2009-12-23 16:11:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl最大消费额 As Double, str收费类别 As String, lng序号 As Long
    Err = 0: On Error GoTo Errhand:
    Call zlInit收费类别Struc
    lng序号 = 0
    With rsFeeList
        .Sort = "收费类别"
        Do While Not rsFeeList.EOF
            If str收费类别 <> Nvl(!收费类别) Then
                grsStatic.rs收费类别汇总.AddNew
                grsStatic.rs收费类别汇总!收费类别 = Nvl(!收费类别)
                str收费类别 = Nvl(!收费类别)
            End If
            grsStatic.rs收费类别汇总!实收金额 = Val(Nvl(grsStatic.rs收费类别汇总!实收金额)) + Val(Nvl(!实收金额))
            grsStatic.rs收费类别汇总.Update
            grsStatic.dbl费用总额 = grsStatic.dbl费用总额 + Val(Nvl(!实收金额))
            
            grsStatic.rs分单类别汇总.Find "分类='" & Nvl(rsFeeList!单据序号) & "_" & Nvl(!收费类别) & "'", , , 1
            If grsStatic.rs分单类别汇总.EOF Then
                grsStatic.rs分单类别汇总.AddNew
                grsStatic.rs分单类别汇总!收费类别 = Nvl(!收费类别)
                
            End If
            grsStatic.rs分单类别汇总!单据序号 = Val(Nvl(!单据序号))
            grsStatic.rs分单类别汇总!实收金额 = Val(Nvl(grsStatic.rs分单类别汇总!实收金额)) + Val(Nvl(!实收金额))
            grsStatic.rs分单类别汇总.Update
            rsFeeList.MoveNext
        Loop
    End With
    zlInit收费类别数据 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zl获取最大消费额(ByVal str限制类别 As String, ByVal dbl最大消费额 As Double, ByVal dbl已刷累计 As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取最大消费额
    '    dbl最大消费额=-1表示未传入最大消费额
    '编制:刘兴洪
    '日期:2009-12-24 10:24:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl限定金额 As Double, dbl可消费 As Double
    Err = 0: On Error GoTo Errhand:
    
    If str限制类别 <> "" Then
        str限制类别 = zlGet获取限制类别FromNameToCode(str限制类别)
    End If
    dbl限定金额 = 0
    If str限制类别 <> "" Then
        With grsStatic.rs收费类别汇总
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                If InStr(1, str限制类别, "," & Nvl(!收费类别) & ",") > 0 Then
                    dbl限定金额 = dbl限定金额 + Val(Nvl(!实收金额))
                End If
                .MoveNext
            Loop
        End With
    End If
    '计算公式:
    '最大可消费额= 总费用-冲预交-已消费额-限定金额
    dbl可消费 = dbl最大消费额 - dbl限定金额 - dbl已刷累计
    zl获取最大消费额 = IIf(dbl可消费 < 0, 0, dbl可消费)
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlGet失效面额(ByVal lng消费卡ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取失效面额
    '返回:失效面额
    '编制:刘兴洪
    '日期:2009-12-23 15:08:04
    '说明：只有在当前时间大于了有效期时才调用该函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, dblTemp As Double
    
    Err = 0: On Error GoTo Errhand:
    strSQL = _
        "Select b.交易序号, Nvl(b.余额, 0) As 失效金额" & vbNewLine & _
        "From 病人卡结算记录 A, 帐户缴款余额 B" & vbNewLine & _
        "Where a.交易序号 = b.交易序号 And a.消费卡id = b.消费卡id And a.记录性质 = 1 And a.消费卡id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取失效额", lng消费卡ID)
    If rsTemp.EOF Then
        '没有记录说明该卡金额已全部使用完
        zlGet失效面额 = 0
        Exit Function
    End If
    
    If Val(Nvl(rsTemp!交易序号)) > 0 Then
        '升级后的发卡记录，直接取失效金额
        dblTemp = Val(Nvl(rsTemp!失效金额))
    Else
        '升级前的发卡记录，需要统计失效金额
        strSQL = _
            "Select Sum(Nvl(失效金额, 0)) As 失效金额" & vbNewLine & _
            "From (" & vbNewLine & _
            "    Select 卡面金额 As 失效金额 From 消费卡信息 A Where ID = [1] And 有效期 < Sysdate" & vbNewLine & _
            "    Union All" & vbNewLine & _
            "    Select Nvl(Sum(a.应收金额), 0) As 失效金额" & vbNewLine & _
            "    From 病人卡结算记录 A, 消费卡信息 B" & vbNewLine & _
            "    Where a.消费卡id = b.Id And a.记录性质 = 4 And a.消费卡id = [1]" & vbNewLine & _
            "          And a.交易时间 <= Nvl(b.有效期, To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
            "     )"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取失效额", lng消费卡ID)
        dblTemp = Val(Nvl(rsTemp!失效金额))
        If dblTemp < 0 Then dblTemp = 0
    End If
    zlGet失效面额 = dblTemp
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlGet获取限制类别FromNameToCode(ByVal str限制类别 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据限制类别名称取相关的编码
    '返回:
    '编制:刘兴洪
    '日期:2009-12-23 16:31:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = zlGet收费类别
    rsTemp.Filter = 0
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    If str限制类别 = "" Then zlGet获取限制类别FromNameToCode = "": Exit Function
    str限制类别 = "," & str限制类别 & ","
    With rsTemp
        Do While Not .EOF
            str限制类别 = Replace(str限制类别, "," & Nvl(rsTemp!名称) & ",", "," & Nvl(rsTemp!编码) & ",")
            .MoveNext
        Loop
    End With
    zlGet获取限制类别FromNameToCode = str限制类别
 End Function
Public Function zl分摊结算数据(ByRef rsRquare As ADODB.Recordset, ByRef rs分摊 As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将刷卡结果分摊结算数据给每张单据明细
    '入出参 rsRquare-(接口编号 消费卡ID,卡号,结算方式,卡名称,余额,结算金额 交易时间,备注,结算标志)
    '       rs分摊-显示每张单据分摊情况
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-06 10:13:43
    '规则说明:
    '   1.先分摊限制类别的
    '   2.再分摊不限制类别的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strTemp As String, str限制类别 As String, dbl金额 As Double
    Dim dbl总额 As Double
    Set rs分摊 = New ADODB.Recordset
    With rs分摊
        If .State = adStateOpen Then .Close
        .Fields.Append "单据序号", adDouble, 18, adFldIsNullable
        .Fields.Append "消费卡ID", adDouble, 18, adFldIsNullable
        .Fields.Append "卡号", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "结算方式", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "分摊额", adDouble, 16, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With

    Set rsTemp = zlDatabase.CopyNewRec(rsRquare)
    Err = 0: On Error GoTo Errhand:
    
    '先确定，存在哪些限制类别
    rsTemp.Filter = "消费卡ID >0"
    str限制类别 = ""
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        strTemp = zlFromCardGet限制类别(Val(Nvl(rsTemp!消费卡ID)), False)
        If InStr(1, str限制类别, strTemp) <= 0 Then
            str限制类别 = str限制类别 & "," & strTemp
        End If
        rsTemp.MoveNext
    Loop
    
    rsTemp.Filter = 0
    If str限制类别 <> "" Then
        str限制类别 = zlGet获取限制类别FromNameToCode(str限制类别) & ","
    End If
    
    rsTemp.Filter = 0
    With grsStatic.rs分单类别汇总
        '先将限制类别的进行分摊
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '需要计算
            If InStr(1, str限制类别, "," & Nvl(!收费类别) & ",") > 0 Then
                '存在限制类别,先将这部分分摊掉
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                   strTemp = zlFromCardGet限制类别(Val(Nvl(rsTemp!消费卡ID)), True)
                   If InStr(1, strTemp, "," & Nvl(!收费类别) & ",") <= 0 And Val(Nvl(rsTemp!结算金额)) > 0 Then
                      '只有用不限定的类别的分摊
                       dbl金额 = Val(Nvl(!实收金额))
                      If dbl金额 >= Val(Nvl(rsTemp!结算金额)) Then
                        dbl金额 = Val(Nvl(rsTemp!结算金额))
                        rsTemp!结算金额 = 0
                        rsTemp.Update
                        !分摊金额 = Val(Nvl(!分摊金额)) + dbl金额
                        .Update
                      Else
                        '小的话
                        rsTemp!结算金额 = Val(Nvl(rsTemp!结算金额)) - dbl金额
                        rsTemp.Update
                        !分摊金额 = Val(Nvl(!分摊金额)) + dbl金额
                      End If
                      rs分摊.Filter = "单据序号=" & Val(Nvl(rsTemp!单据序号)) & " And 消费卡ID=" & Val(Nvl(rsTemp!消费卡ID)) & " And 卡号='" & Nvl(rsTemp!卡号) & "'"
                      If rs分摊.EOF Then
                          rs分摊.AddNew
                      End If
                      rs分摊!单据序号 = Val(Nvl(rsTemp!单据序号))
                      rs分摊!消费卡ID = Val(Nvl(rsTemp!消费卡ID))
                      rs分摊!卡号 = Nvl(rsTemp!卡号)
                      rs分摊!结算方式 = Trim(Nvl(rsTemp!结算方式))
                      rs分摊!分摊额 = Val(Nvl(rs分摊!分摊额)) + dbl金额
                      rs分摊.Update
                   End If
                   If !分摊金额 = !实收金额 Then Exit Do
                   rsTemp.MoveNext
                Loop
            End If
            .MoveNext
        Loop
        '再分摊不限定的
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
             If Val(Nvl(!分摊金额)) <= Val(Nvl(!实收金额)) Then
                
                rsTemp.Filter = 0
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                   strTemp = zlFromCardGet限制类别(Val(Nvl(rsTemp!消费卡ID)), True)
                   If InStr(1, strTemp, "," & Nvl(!收费类别) & ",") <= 0 And Val(Nvl(rsTemp!结算金额)) > 0 Then
                      dbl金额 = Val(Nvl(!实收金额))
                      If dbl金额 >= Val(Nvl(rsTemp!结算金额)) Then
                        dbl金额 = Val(Nvl(rsTemp!结算金额))
                        rsTemp!结算金额 = 0
                        rsTemp.Update
                        !分摊金额 = Val(Nvl(!分摊金额)) + dbl金额
                        .Update
                      Else
                        '小的话
                        rsTemp!结算金额 = Val(Nvl(rsTemp!结算金额)) - dbl金额
                        rsTemp.Update
                        !分摊金额 = Val(Nvl(!分摊金额)) + dbl金额
                      End If
                      rs分摊.Filter = "单据序号=" & Val(Nvl(!单据序号)) & " And 消费卡ID=" & Val(Nvl(rsTemp!消费卡ID)) & " And 卡号='" & Nvl(rsTemp!卡号) & "'"
                      If rs分摊.EOF Then
                          rs分摊.AddNew
                      End If
                      rs分摊!单据序号 = Val(Nvl(!单据序号))
                      rs分摊!消费卡ID = Val(Nvl(rsTemp!消费卡ID))
                      rs分摊!卡号 = Nvl(rsTemp!卡号)
                      rs分摊!结算方式 = Trim(Nvl(rsTemp!结算方式))
                      rs分摊!分摊额 = Val(Nvl(rs分摊!分摊额)) + dbl金额
                      rs分摊.Update
                   End If
                   If !分摊金额 = !实收金额 Then Exit Do
                   rsTemp.MoveNext
                Loop
             End If
             .MoveNext
        Loop
    End With
    
    With rs分摊
        .Filter = 0
        If .RecordCount > 0 Then .MoveFirst
        dbl金额 = 0
        Do While Not .EOF
            dbl金额 = dbl金额 + Val(Nvl(!分摊额))
            .MoveNext
        Loop
    End With
    dbl总额 = 0
    With rsRquare
        .Filter = 0
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            dbl总额 = dbl总额 + Val(Nvl(!结算金额))
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    
    If Round(dbl总额, 4) <> Round(dbl金额, 4) Then
        ShowMsgbox "多单据分摊时，出现了不等情况,请重新刷卡!"
        Exit Function
    End If
    '检查计算后的明细分摊额与总的是否一致
    zl分摊结算数据 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlFromCardGet限制类别(ByVal lng消费卡ID As Long, ByVal blnCode As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据消费卡,获取相关的限定类另
    '入参:lng消费卡ID-消费卡ID
    '     blnCode-编码
    '出参:
    '返回:返回限制类别串
    '编制:刘兴洪
    '日期:2010-01-06 11:18:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, str限制类别 As String
    Err = 0: On Error GoTo Errhand:
    gstrSQL = "Select 限制类别 From 消费卡信息 Where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取消费卡信息的限制类别", lng消费卡ID)
    If rsTemp.EOF Then Exit Function
    str限制类别 = Nvl(rsTemp!限制类别)
    If blnCode Then
        zlFromCardGet限制类别 = zlGet获取限制类别FromNameToCode(str限制类别)
    Else
        zlFromCardGet限制类别 = str限制类别
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlGetRquare(ByVal str结帐ID_IN As String, ByRef rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡结算交易时的相关预结数据
    '入参:str结帐ID_IN-指定的结算ID
    '出参:rsSquare-结帐数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-15 11:08:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, lngID As Long
    
    On Error GoTo errHandle
    
    Call zlInitBrushCardRec(rsSquare)
    If str结帐ID_IN = "" Then str结帐ID_IN = "0"
    
    strSQL = _
        "Select /*+ cardinality(j,10)*/ Distinct a.Id, 接口编号, a.消费卡id, a.序号, a.记录状态, a.结算方式," & vbNewLine & _
        "      a.应收金额 As 结算金额, a.卡号, a.交易流水号, a.交易时间, a.备注, a.结算标志, c.结帐id" & vbNewLine & _
        "From 病人预交记录 C, 病人卡结算记录 A," & vbNewLine & _
        "     (Select Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) J" & vbNewLine & _
        "Where a.结算id = c.Id And c.结帐id = j.Column_Value And a.结算标志 = 0 And c.记录状态 = 1" & vbNewLine & _
        "Order By ID, 结帐id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取结帐ID的相关刷卡信息", str结帐ID_IN)
    gTy_TestBug.bln补调交易 = True
    With rsSquare
        Do While Not rsTemp.EOF
            If lngID <> Val(Nvl(rsTemp!id)) Then
                .AddNew
                !接口编号 = Val(Nvl(rsTemp!接口编号))
                !消费卡ID = Val(Nvl(rsTemp!消费卡ID))
                !卡号 = Nvl(rsTemp!卡号)
                !结算方式 = Nvl(rsTemp!结算方式)
                !卡名称 = zlGet接口名称(Val(Nvl(rsTemp!接口编号)))
                !余额 = 0
                !结算金额 = Val(Nvl(rsTemp!结算金额))
                !交易时间 = rsTemp!交易时间
                !交易流水号 = IIf(Val(Nvl(rsTemp!消费卡ID)) = 0, Nvl(rsTemp!交易流水号), Nvl(rsTemp!id))     '对于，消费卡的处理，没有特别的处理，在补传交易时，只是模拟作用。简单的更新相关的标识
                !备注 = Nvl(rsTemp!备注)
                !结算标志 = 0
            End If
            !分摊页码 = Nvl(!分摊页码) & "," & Val(Nvl(rsTemp!结帐ID))
            .Update
            rsTemp.MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    zlGetRquare = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlGet接口名称(ByVal lng接口编号 As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取接口名称
    '返回:接口名称
    '编制:刘兴洪
    '日期:2010-01-15 11:23:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp  As ADODB.Recordset
    Set rsTemp = zlGet消费卡接口
    rsTemp.Filter = "编号=" & lng接口编号
    If rsTemp.EOF Then
        zlGet接口名称 = ""
    Else
        zlGet接口名称 = Nvl(rsTemp!名称)
    End If
End Function
Public Function zlGet接口编号(ByVal lng预交ID As Long) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '功能：根据预交ID,获取相应的接口编号
    '返回:结算卡的接口ID
    '编制：刘兴洪
    '日期：2010-06-18 14:05:08
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = " Select distinct A.接口编号 From  病人卡结算记录 A Where A.结算ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取退单的接口编号", lng预交ID)
    If rsTemp.RecordCount = 0 Then zlGet接口编号 = 0: Exit Function
    zlGet接口编号 = Val(Nvl(rsTemp!接口编号))
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlSave卡结算记录(ByVal lng预交ID As Long, ByVal strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：保存相关的结算数据
    '           用||分隔: 接口编号||消费卡ID(可传'')||结算方式||结算金额||卡号||交易流水号||交易时间(yyyy-mm-dd hh24:mi:ss)||备注
    '编制：刘兴洪
    '日期：2010-06-18 16:07:05
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, strSQL As String, strTemp As String
    
    If strBlanceInfor = "" Then Exit Function
    varData = Split(strBlanceInfor, "||")
    If UBound(varData) < 7 Then Exit Function
    
    'Zl_病人卡结算记录_支付
    strSQL = "Zl_病人卡结算记录_支付("
    '  接口编号_In   消费卡类别目录.编号%Type,
    strSQL = strSQL & "" & Val(varData(0)) & ","
    '  卡号_In       消费卡信息.卡号%Type,
    strSQL = strSQL & "'" & Trim(varData(4)) & "',"
    '  消费卡id_In   消费卡信息.Id%Type,
    strSQL = strSQL & "" & Val(varData(1)) & ","
    '  结算金额_In   病人卡结算记录.应收金额%Type,
    strSQL = strSQL & "" & Val(varData(3)) & ","
    '  预交id_In     病人预交记录.Id%Type,
    strSQL = strSQL & "" & lng预交ID & ","
    '  操作员编号_In 病人卡结算记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人卡结算记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  收款时间_In   病人预交记录.收款时间%Type
    If Trim(varData(6)) = "" Or IsDate(varData(6)) = False Then
        strTemp = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        strTemp = Trim(varData(6))
    End If
    If strTemp = "" Then
        strSQL = strSQL & "NULL)"
    Else
        strSQL = strSQL & "to_date('" & strTemp & "','yyyy-mm-dd hh24:mi:ss'))"
    End If
    zlDatabase.ExecuteProcedure strSQL, "保存卡结算记录"
    zlSave卡结算记录 = True
End Function

Public Function zlInputIsCard(ByRef txtInput As Object, ByVal KeyAscii As Integer, ByVal lngSys As Long, Optional ByVal blnPassWd As Boolean = False) As Boolean
'功能：判断指定文本框中当前输入是否在刷卡(是否达到卡号长度，在调用程序中判断),并根据系统参数处理是否密文显示
'参数：KeyAscii=在KeyPress事件中调用的参数
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, strText As String
    
     '刷卡时含有特殊符号的由调用方取消输入
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then Exit Function
    
    '处理当前键入后显示的内容(还未显示出来)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    '判断是否在刷卡
    If KeyAscii > 32 Then
        sngNow = timer
        If txtInput.Text = "" Or strText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True   '用一台笔记本测试，一般在0.014左右
        End If
    End If
    '刷卡时卡号是否密文显示
    If blnCard Then
        txtInput.PasswordChar = IIf(Not blnPassWd, "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    zlInputIsCard = blnCard
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
        str预约方式 = rsTemp!预约方式
        rsTemp.MoveNext
    Wend
    zl_Get预约方式ByNo = str预约方式
End Function

Public Sub CreateSquareCardObject(ByRef frmMain As Object, _
    ByVal lngModule As Long, Optional cnOracle As ADODB.Connection)
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
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, IIf(cnOracle Is Nothing, gcnOracle, cnOracle), False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Sub
    End If
End Sub
Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         'Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub

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
    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
End Sub

Public Function zlGet支付方式(ByVal lng卡类别ID As Long, ByVal str结算方式 As String) As String
    '根据结算方式查找支付方式
    Dim strSQL As String, rsTemp As Recordset
    '名称|结算方式|是否退现|是否全退|结算性质
    zlGet支付方式 = str结算方式 & "|" & str结算方式 & "|1|0"
    On Error GoTo Errhand
    strSQL = "" & _
            " Select A.名称,A.是否退现,A.是否全退,B.性质 from 医疗卡类别 A,结算方式 B where A.结算方式 = B.名称 And A.ID = [1] And A.结算方式=[2]" & _
            " Union All " & _
            " Select A.名称,A.是否退现,A.是否全退,B.性质 from 消费卡类别目录 A,结算方式 B where A.结算方式 = B.名称 And A.编号=[1] And A.结算方式=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取支付卡结算方式", lng卡类别ID, str结算方式)
    If Not rsTemp.EOF Then
        zlGet支付方式 = Nvl(rsTemp!名称, str结算方式) & "|" & str结算方式 & "|" & Nvl(rsTemp!是否退现, 1) & "|" & Nvl(rsTemp!是否全退, 0) & "|" & Nvl(rsTemp!性质, 0)
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlFormatNum(ByVal dblMoney As Double) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取格式化串(比如:.03 格式为0.03,123格式为123)
    '入参:dblMoney-格式化金额
    '返回:返回格式化串(比如:.03 格式为0.03,123格式为123)
    '编制:刘兴洪
    '日期:2014-07-30 15:29:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    Dim strMoney As String
'    If dblMoney = 0 Then Exit Function
    strTemp = Format(dblMoney, "###0.00######;-###0.00######;;")
    If strTemp = "" Then Exit Function
    strMoney = strTemp
    For i = Len(strTemp) To 1 Step -1
        If Val(Mid(strTemp, i, 1)) <> 0 Or Mid(strTemp, i, 1) = "." Then Exit For
        strMoney = Mid(strTemp, 1, i - 1)
    Next
    If Right(strMoney, 1) = "." Then strMoney = Mid(strMoney, 1, Len(strMoney) - 1)
    zlFormatNum = strMoney
End Function

Public Sub SetEnabledBackColor(ByVal frmMain As Form)
    '功能:设置窗口中所有控件可用状态与不可用状态的背景颜色
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To frmMain.Controls.count - 1
        If UCase(TypeName(frmMain.Controls(i))) = UCase("TextBox") _
            Or UCase(TypeName(frmMain.Controls(i))) = UCase("ComboBox") Then
            Call zl_SetCtlBackColor(frmMain.Controls(i), frmMain)
        End If
    Next
End Sub

Public Function Ceil(ByVal dblNum As Double) As Integer
    '功能:向上取整
    If dblNum > 0 Then
        Ceil = -1 * Int(-1 * dblNum)
    Else
        Ceil = Fix(dblNum)
    End If
End Function

Public Function Floor(ByVal dblNum As Double) As Integer
    '功能:向下取整
    If dblNum > 0 Then
        Floor = -1 * Fix(-1 * dblNum)
    Else
        Floor = Int(dblNum)
    End If
End Function

Public Function FromStringListBulidSQL(ByVal bytBulidType As Byte, ByVal strValues As String, _
    ByRef varPara As Variant, ByRef strBulitSQL As String, _
    ByVal strColumnAliaName As String, Optional intStartPara As Integer = 1) As Boolean
    '功能:将参数值(值列表组成的)超长的参数分解为含有多个参数的SQL,如:select ... From str2List Union ALL Selelct ..
    '入参:strValues-值,多个用逗号分离
    '     strColumnAliaName-列别名
    '     bytType-0-字符型;1-数字型;
    '     intStartPara-启动的参数序号
    '出参:varPara-返回的参数值数据组
    '     strBulitSQL-返回的构建的SQL串
    '返回:如果获取成功,返回true,否则返回False
    Dim varData As Variant, strTemp As String
    Dim i As Long, j As Long, strSQL As String
    Dim strTable As String, strColumnName As String
    
    On Error GoTo ErrHandler
    strColumnName = " a.Column_Value "
    If strColumnAliaName <> "" Then strColumnName = strColumnName & " As " & strColumnAliaName
    
    If bytBulidType = 0 Then
        strTable = "Table(f_str2list([0]))"
    Else
        strTable = "Table(f_Num2list([0]))"
    End If
    
    j = intStartPara
    ReDim Preserve varPara(0 To j - 1)
    
    varData = Split(strValues, ",")
    strTemp = ""
    For i = 0 To UBound(varData)
        If zlCommFun.ActualLen(strTemp & "," & varData(i)) > 4000 Then
            strSQL = strSQL & " Union ALL " & _
                " Select /*+cardinality(a,10) */" & strColumnName & _
                " From " & Replace(strTable, "[0]", "[" & j & "]") & " A"
            ReDim Preserve varPara(0 To j - 1)
            varPara(j - 1) = Mid(strTemp, 2)
            j = j + 1: strTemp = ""
        End If
        strTemp = strTemp & "," & varData(i)
    Next
    If strTemp <> "" Then
        strSQL = strSQL & " Union ALL " & _
            " Select /*+cardinality(a,10) */" & strColumnName & _
            " From " & Replace(strTable, "[0]", "[" & j & "]") & " A"
        ReDim Preserve varPara(0 To j - 1)
        varPara(j - 1) = Mid(strTemp, 2)
    End If
    
    If strSQL <> "" Then strSQL = Mid(strSQL, 11)
    strBulitSQL = strSQL
    FromStringListBulidSQL = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SplitCardNos(ByVal strCardNoRange As String, ByRef strCardNos As String) As Boolean
    '功能:根据传入的卡号范围，分解成相关的卡号
    '入参:
    '   strCardNoRange-卡号范围
    '出参:
    '   strCardNos-返回卡号数(用逗号分离)
    '返回:分解成功返回True，否则返回False
    Dim varData As Variant, lngCount As Long
    Dim strCardStartNO As String, strCardEndNO As String, strCurNo As String
    Dim str数量 As String

    varData = Split(strCardNoRange & "～", "～")
    strCardStartNO = varData(0): strCardEndNO = varData(1)
    If strCardEndNO = "" Then
        strCardNos = strCardStartNO
        SplitCardNos = True
        Exit Function
    End If
    If strCardStartNO > strCardEndNO Then Exit Function
    
    str数量 = zlstr.ExpressValue(strCardEndNO & "-" & strCardStartNO & "+1")
    If InStr(UCase(str数量), "E") > 0 Or Len(str数量) > 4 Then '数量太大已经变成科学计算法
        ShowMsgbox "卡号范围不能大于10000，请分段发放！"
        Exit Function
    End If
    
    strCurNo = strCardStartNO
    strCardNos = strCardStartNO
    Do While True
        If strCurNo >= strCardEndNO Then Exit Do
        strCurNo = zlstr.Increase(strCurNo)
        strCardNos = strCardNos & "," & strCurNo
        
        lngCount = lngCount + 1
        If lngCount > 10000 Then
            ShowMsgbox "卡号范围不能大于10000，请分段发放！"
            Exit Function
        End If
    Loop
    SplitCardNos = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CollExitsValue(ByVal coll As Collection, ByVal strKey As String) As Boolean
'功能：根据关键字判断元素是否存在于集合中
    Dim blnExits As Boolean
    
    If coll Is Nothing Then Exit Function
    CollExitsValue = True
    Err = 0: On Error Resume Next
    blnExits = IsObject(coll(strKey))
    If Err <> 0 Then Err = 0: CollExitsValue = False
End Function

Public Sub CheckInputPassWord(KeyAscii As Integer, Optional ByVal blnOnlyNum As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查密码输入
    '编制:刘兴洪
    '日期:2011-07-07 00:40:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 8 Or KeyAscii = 13 Then Exit Sub
    If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If blnOnlyNum Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
       If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
            If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                 If InStr(1, "!@#$%^&*()_+-=><?,:;~`./", Asc(KeyAscii)) = 0 Then KeyAscii = 0
            End If
       End If
    End If
End Sub

Private Sub ClearYLCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除相关的医疗卡卡对象
    '编制:刘兴洪
    '日期:2018-02-13 11:45:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errH
    If Not gObjYLCardObjs Is Nothing Then
        For i = 1 To gObjYLCardObjs.count
            If Not gObjYLCardObjs(i).CardObject Is Nothing Then
                Call gObjYLCardObjs(i).CardObject.zlReleaseComponent
            End If
            Set gObjYLCardObjs(i).CardObject = Nothing
            gObjYLCardObjs(i).InitCompents = False
        Next
    End If
    Set gObjYLCardObjs = Nothing
    Exit Sub
errH:
    Resume Next
End Sub

 Public Sub zlCloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub

Public Function zlCloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    zlCloseWindows = (Forms.count = 0)
End Function

Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:释放资源
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-13 10:30:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '实例数为0时，才放资源
    If glngInstanceCount > 0 Then Exit Function
    
    Call ClearYLCardObject
 
   
    Call zlCloseSquareCardObject '释放结算卡相关资源
    Call zlCloseWindows   '关闭窗体
    
    Err = 0: On Error Resume Next
    
    If Not gobjComLib Is Nothing Then Set gobjComLib = Nothing
    If Not gobjCommFun Is Nothing Then Set gobjCommFun = Nothing
    If Not gobjControl Is Nothing Then Set gobjControl = Nothing
    If Not gobjDataBase Is Nothing Then Set gobjDataBase = Nothing
    If Not gobjPublicExpense Is Nothing Then Set gobjPublicExpense = Nothing
    If Not gobjStartCards Is Nothing Then Set gobjStartCards = Nothing
    If Not gObjYLCards Is Nothing Then Set gObjYLCards = Nothing
    If Not gcolPrivs Is Nothing Then Set gcolPrivs = Nothing
    If Not gfrmMain Is Nothing Then Set gfrmMain = Nothing
    If Not gfrmCardMgr Is Nothing Then Set gfrmCardMgr = Nothing
    If Not gobjXml Is Nothing Then Set gobjXml = Nothing
    If Not gObjXFCards Is Nothing Then Set gObjXFCards = Nothing
    If Not grs医疗卡类别 Is Nothing Then Set grs医疗卡类别 = Nothing
        If Not grsStatic.rs收费类别 Is Nothing Then
        If grsStatic.rs收费类别.State = 1 Then grsStatic.rs收费类别.Close
    End If
    If Not grsStatic.rs消费卡接口 Is Nothing Then
        If grsStatic.rs消费卡接口.State = 1 Then grsStatic.rs消费卡接口.Close
    End If
    
    If Not grs医疗卡类别 Is Nothing Then Set grs医疗卡类别 = Nothing
    If Not grsStatic.rs收费类别 Is Nothing Then Set grsStatic.rs收费类别 = Nothing
    If Not grsStatic.rs消费卡接口 Is Nothing Then Set grsStatic.rs消费卡接口 = Nothing
    If Not grsStatic.rs分单类别汇总 Is Nothing Then Set grsStatic.rs分单类别汇总 = Nothing
    If Not grsStatic.rs收费类别汇总 Is Nothing Then Set grsStatic.rs收费类别汇总 = Nothing
    If Not grsSystem Is Nothing Then Set grsSystem = Nothing
    If Not grsOneCard Is Nothing Then Set grsOneCard = Nothing
    If Not grs医疗付款方式 Is Nothing Then Set grs医疗付款方式 = Nothing
    zlReleaseResources = True
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


