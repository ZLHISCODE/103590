Attribute VB_Name = "mdl重庆"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)

'API函数声明

'1、接口初始化：检查整个网络环境是否畅通，包括医院客户端与前置机、前置机与中心服务器间。
Private Declare Function dy_Init Lib "SiInterface" Alias "INIT" () As Long

'2 业务处理：调用执行医保业务所需要的处理
Private Declare Function dy_Business_Handle Lib "SiInterface" Alias "BUSINESS_HANDLE" _
    (ByVal InputData As String, ByVal OutputData As String) As Long

'Private gobj重庆市东软 As New clsT_CQDRYB
Private mstr费用终止日期 As String                  '保存本次住院结算操作员选择的费用终止日期
Private mstr医保号 As String
Private mdbl余额 As Double
Private mlng病人ID As Long
Private mstr门诊号 As String
Private mstr中心编号 As String

Private mblnIint As Boolean
Private mblnFail As Boolean                         '初始化失败脱机处理

Private gstr就诊时间 As String
Private mdbl费用总额 As Double
Private mbln多单据收费 As Boolean
Private mbln结算 As Boolean                         '多单据下门诊只调用医保一次结算，用此做判断
Private mstr结算流水号 As String                    '用于多单据收费的每条记录都保存相同的流水号，便于查证
Private mcnYB As New ADODB.Connection

'以下结构体用于纪录虚拟结算结果，以便在结算时核对
Private Type typBalance
    cur个人帐户 As Double
    cur统筹支付 As Double
    cur大额统筹 As Double
    cur公务员补助 As Double
    cur公务员返还 As Double
    cur医院收支 As Double
    cur民政补助 As Double
    HIS计算医院收支 As Boolean          '是否由HIS来计算医院收支
End Type
Private pre_Balance As typBalance

'###############################################################################
'20061113,zyb:取消TrackRecordInsure()的调用，避免打开医保前置机连接数过多（东软限制了连接数，调dy_init也会打开一个连接）
'###############################################################################

Public Function 医保初始化_重庆() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    Dim lngReturn As Long
    
    If mblnIint = True Then
        '只需要调用一次
        医保初始化_重庆 = True
        Exit Function
    End If
    
    On Error Resume Next
    
    If gclsInsure.GetCapability(Support初始化失败脱机处理, 0, TYPE_重庆市) And mblnFail Then
        Exit Function
    Else
        lngReturn = dy_Init
'        lngReturn = gobj重庆市东软.dy_Init
    End If
    If Err <> 0 Then
        MsgBox "不能正确调用医保接口程序。", vbInformation, gstrSysName
        mblnFail = True
        Exit Function
    End If
    
    If lngReturn = -1 Then
        mblnFail = True
        MsgBox "不能完成医保接口初始化工作，请检查整个网络环境是否畅通。包括：" & vbCrLf & vbCrLf & _
          "1、医院客户端与医院前置机应用服务器之间；" & vbCrLf & _
          "2、医院前置机应用服务器与医保中心应用服务器之间。", vbInformation, gstrSysName
    Else
        医保初始化_重庆 = True
        mblnIint = True
    End If
End Function

Public Function 身份标识_重庆(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    'alter table 保险帐户 add (封锁信息 varchar2(200),病种名称 varchar2(250));
    Dim str医保号 As String, StrInput As String, arrOutput  As Variant, int类别 As Integer
    Dim STR姓名 As String, str性别 As String, str身份证号码 As String, lng年龄 As Long
    Dim str出生日期 As String, str人员类别 As String, str单位编码 As String, str单位名称 As String
    Dim strIdentify As String, str附加 As String, str门诊号 As String, str封锁信息 As String
    Dim datCurr As Date
    Dim lngTemp As Long, str上次就诊时间 As String
    Dim bln离休就诊录入诊断 As Boolean
    Dim int业务类型 As Integer
    Dim str分类 As String, str疾病编码 As String, str并发症 As String, str疾病名称 As String
    Dim rsTemp As New ADODB.Recordset
    Dim rs病种 As ADODB.Recordset
    
    Call DebugTool("进入身份验证")
    '初始化一些变量
    mlng病人ID = 0
    mstr门诊号 = ""
    mstr医保号 = ""
    mdbl余额 = 0
    
    '如果是挂号，则直接当做门诊处理
    If bytType = 3 Then bytType = 0
    int类别 = bytType
    If frmIdentify重庆.GetIdentify(str医保号, int类别) = False Then
        Exit Function
    End If
    
    Call DebugTool("关闭身份验证窗体")
    '取保险参数“离休就诊录入诊断”
    gstrSQL = "Select Nvl(参数值,0) AS 参数值 From 保险参数 Where 险类=[1] And 参数名='离休就诊录入诊断'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险参数“离休就诊录入诊断”", TYPE_重庆市)
    If rsTemp.RecordCount <> 0 Then
        bln离休就诊录入诊断 = rsTemp!参数值
    End If
    Call DebugTool("取保险参数")
    
    '如果当天已进行过收费，则不必再次刷卡，就诊流水号以上次的号为准
    If bytType = 0 Then lngTemp = GetRegisted(str医保号, str上次就诊时间)
    datCurr = zlDatabase.Currentdate
    Call DebugTool("判断当天是否已收过费")
    
    '虽然当天就诊过，但程序后面仍然是产生了新的流水号并进行医保登记
    If lngTemp <> 0 Then
        Call DebugTool("lngTemp<>0")
        
        lng病人ID = lngTemp
        '提取上次就诊的业务类型
        gstrSQL = "Select 业务类型,帐户余额,退休证号,病种名称,并发症 From 保险帐户 Where 险类=[1] ANd 病人ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取业务类型", TYPE_重庆市, lng病人ID)
        int业务类型 = rsTemp!业务类型
        mdbl余额 = Nvl(rsTemp!帐户余额, 0)
        str疾病编码 = Nvl(rsTemp!退休证号)
        str疾病名称 = Nvl(rsTemp!病种名称)
        str并发症 = Nvl(rsTemp!并发症)
        '提取病人信息（一天虽产生多条就诊登记记录，但其医保流水号是不变的，因此随便取一条）
        gstrSQL = " Select A.姓名,A.性别,A.年龄,A.身份证号,A.出生日期,B.人员身份 AS 人员类别,A.工作单位 AS 单位名称,B.业务类型,B.中心编号,C.HIS流水号 AS 顺序号 " & _
                  " From 病人信息 A,保险帐户 B,就诊登记记录 C" & _
                  " Where C.主页ID=0 And C.记录ID Is Not NULL And C.病人ID=A.病人ID And A.病人ID=B.病人ID And B.险类=C.险类" & _
                  " And C.险类=[1] And C.就诊时间=[2] And C.病人ID=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", TYPE_重庆市, CDate(str上次就诊时间), lng病人ID)
        
        STR姓名 = rsTemp!姓名
        str性别 = rsTemp!性别
        lng年龄 = Val(Nvl(rsTemp!年龄, 0))
        str身份证号码 = Nvl(rsTemp!身份证号)
        str出生日期 = Format(rsTemp!出生日期, "yyyy-MM-dd")
        
        str人员类别 = Nvl(rsTemp!人员类别)
        str单位编码 = ""
        str单位名称 = Nvl(rsTemp!单位名称) '50的长度，还要扣除2个符号
        mstr中心编号 = Nvl(rsTemp!中心编号)
        str门诊号 = Nvl(rsTemp!顺序号)
        
        mlng病人ID = lng病人ID
        mstr门诊号 = str门诊号
        mstr医保号 = str医保号
        
        Call DebugTool("已成功提取病人信息")
    Else
        Call DebugTool("准备调用01交易")
        '调用接口
        StrInput = "01|" & str医保号
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
        
        '取得返回值
        STR姓名 = arrOutput(1)
        str性别 = arrOutput(2)
        lng年龄 = Val(arrOutput(3))
        str身份证号码 = arrOutput(4)
        str出生日期 = Get出生日期(str身份证号码, lng年龄)
        
        str人员类别 = ToVarchar(arrOutput(7), 8) 'VARCHAR2 (20) 在职，在职驻外，临时用工，自由职业军转干，退休，退休异地居住，退职，退职异地居住等
        'arrOutput(8)   公务员标志               'VARCHAR2 (3)
        str单位编码 = ""
        str单位名称 = ToVarchar(arrOutput(9), 48) '50的长度，还要扣除2个符号
        mstr中心编号 = arrOutput(10)
        
        If arrOutput(11) = "2" Then
            MsgBox "该病人医保卡不能继续使用。" & arrOutput(12)
            Exit Function
        End If
        
        str封锁信息 = ""
        If arrOutput(11) <> "0" Then
            '住院时要提醒
            str封锁信息 = arrOutput(12)
            MsgBox str封锁信息, vbInformation, gstrSysName
        End If
        Call DebugTool("01交易调用成功！")
    End If
    
    '卡号;医保号;密码;姓名;性别;出生日期;身份证;工作单位
    '医保号第一位为卡类型
    strIdentify = str医保号 & ";" & str医保号 & ";;" & STR姓名 & ";" & str性别 & ";" & str出生日期 & ";" & str身份证号码 & ";" & str单位名称 & "(" & str单位编码 & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    str附加 = ";"                                       '8.中心代码
    str附加 = str附加 & ";"                             '9.顺序号
    str附加 = str附加 & ";" & str人员类别               '10人员身份
    str附加 = str附加 & ";0"                            '11帐户余额
    str附加 = str附加 & ";0"                            '12当前状态
    str附加 = str附加 & ";"                             '13病种ID
    str附加 = str附加 & ";" & IIf(Left(str人员类别, 1) = "退", 2, 1)     '14在职(1,2)
    str附加 = str附加 & ";"                             '15退休证号
    str附加 = str附加 & ";" & lng年龄                   '16年龄段
    str附加 = str附加 & ";"                             '17灰度级
    str附加 = str附加 & ";0"                            '18帐户增加累计
    str附加 = str附加 & ";0"                            '19帐户支出累计
    str附加 = str附加 & ";"                             '20进入统筹累计
    str附加 = str附加 & ";"                             '21统筹报销累计
    str附加 = str附加 & ";"                             '22住院次数累计
    str附加 = str附加 & ";" & IIf(int类别 = 14, 1, "")  '23就诊类型 (1、急诊门诊)
    
'    If lngTemp = 0 Then
        '如果是第一次就诊
        lng病人ID = BuildPatiInfo(bytType, strIdentify & str附加, lng病人ID, TYPE_重庆市)
        str门诊号 = ToVarchar(lng病人ID & Format(datCurr, "yyMMddHHmmss"), 18)
'    End If
    Call DebugTool("成功产生病人档案")
    
    If bytType = 0 Then        '如果是门诊，同时进行就诊登记
'        '如果是特殊病或急诊抢救，需要选择病人疾病
'        If lngTemp <> 0 Then
'            '当天就诊过
'            If int业务类型 <> int类别 Or int类别 = 13 Or int类别 = 14 Or Mid(str医保号, 1, 1) = "2" Then
'                If int类别 = 13 Or int类别 = 14 Or Mid(str医保号, 1, 1) = "2" Then
'                    If int类别 = 13 Then
'                        '获得审批信息
'                        strInput = "07|" & str医保号
'                        If HandleBusiness(strInput, arrOutput) = False Then Exit Function
'
'                        str分类 = "特殊病"
'                        If frm疾病选择重庆.GetCode(arrOutput, str分类, str疾病编码, str疾病名称, str并发症) = False Then Exit Function
'                    ElseIf int类别 = 14 Then
'                        str分类 = "急诊"
'                        If frm疾病选择重庆.GetCode("", str分类, str疾病编码, str疾病名称, str并发症) = False Then Exit Function
'                    Else
'                        str分类 = "出院"
'                        If frm疾病选择重庆.GetCode("", str分类, str疾病编码, str疾病名称, str并发症) = False Then Exit Function
'                    End If
'                Else
'                    str疾病编码 = ""
'                    str疾病名称 = ""
'                    str并发症 = ""
'                End If
'
'                '需要调用就诊信息修改，才能正常进行结算
'                '调用接口更新：住院号(门诊号)|更新标志|医疗类别|科室|医生|入院日期|入院诊断|出院日期|确诊疾病编码
'                              '|出院原因|经办人|并发症
'                strInput = "03|" & str门诊号 & "|100000" & IIf(str疾病编码 <> "", "1", "0") & "00" & IIf(str并发症 <> "", "1", "0") & "|" & int类别 & "||||||" & str疾病编码 & "|||" & str并发症
'                If HandleBusiness(strInput, arrOutput) = False Then Exit Function
'
'                gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'退休证号','''" & str疾病编码 & "''')"
'                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
'                gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'病种名称','''" & str疾病名称 & "''')"
'                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
'                gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'并发症','''" & str并发症 & "''')"
'                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新并发症")
'            End If
'        Else
    
            Call DebugTool("准备弹出病种选择窗体")
            If int类别 = 13 Or int类别 = 15 Or int类别 = 14 Or (mstr中心编号 = "20" And bln离休就诊录入诊断) Then
                If int类别 = 13 Then
                    '获得审批信息
                    StrInput = "07|" & str医保号
                    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
                    
                    str分类 = "特殊病"
                    If frm疾病选择重庆.GetCode(arrOutput, str分类, str疾病编码, str疾病名称, str并发症) = False Then Exit Function
                ElseIf int类别 = 14 Then
                    str分类 = "急诊"
                    If frm疾病选择重庆.GetCode("", str分类, str疾病编码, str疾病名称, str并发症) = False Then Exit Function
                Else
                    str分类 = "出院"
                    If frm疾病选择重庆.GetCode("", str分类, str疾病编码, str疾病名称, str并发症) = False Then Exit Function
                End If
                
                gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'退休证号','''" & str疾病编码 & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
                gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'病种名称','''" & str疾病名称 & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
                gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'并发症','''" & str并发症 & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新并发症")
            End If
            
            Call DebugTool("准备调用就诊登记")
            StrInput = "02|" & str门诊号 & "|" & int类别 & "|" & str医保号 & _
                       "|门诊|" & ToVarchar(UserInfo.姓名, 20) & "|" & _
                       Format(datCurr, "yyyy-MM-dd") & "|" & str疾病编码 & "|" & ToVarchar(UserInfo.姓名, 20) & "|" & str并发症
            If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
            
            mlng病人ID = lng病人ID
            mstr门诊号 = str门诊号
            mstr医保号 = str医保号
            mdbl余额 = Val(arrOutput(2))
'        End If
    End If
     
     'Modified by ZYB 2006-02-6
    '调用接口更新：住院号(门诊号)|更新标志|医疗类别|科室|医生|入院日期|入院诊断|出院日期|确诊疾病编码
                  '|出院原因|经办人|并发症
    If int类别 = 15 Or int类别 = 14 Then
        Call DebugTool("准备调用03交易")
        StrInput = "03|" & str门诊号 & "|0000001001||" & _
            "|||||" & str疾病编码 & "|||" & str并发症
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    End If
    
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_重庆市 & "," & Year(datCurr) & "," & _
        mdbl余额 & ",0,0,0,0,0,0,0,0,0,'')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    If lngTemp = 0 Then
        '更新封锁信息
        gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'封锁信息','''" & str封锁信息 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
        
        '就诊号不更新到保险帐户中，产生就诊登记记录
        '反正住院都是要更新病种名称、并发症等信息才能结算的，因此这些内容仍然保存在保险帐户中，就诊登记记录只是附带着填写
'        gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'顺序号','''" & str门诊号 & "''')"
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
        
    End If
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'中心编号','''" & mstr中心编号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新中心编号")
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'高龄病床','''" & "0" & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "取消高龄病床标志")
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'业务类型','''" & int类别 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
    
    '就诊号不更新到保险帐户中，产生就诊登记记录
    '反正住院都是要更新病种名称、并发症等信息才能结算的，因此这些内容仍然保存在保险帐户中，就诊登记记录只是附带着填写
    '填写就诊登记记录，参数如下：
    '险类_IN             就诊登记记录.险类%TYPE,
    '病人ID_IN           就诊登记记录.病人ID%TYPE,
    '主页ID_IN           就诊登记记录.主页ID%TYPE,
    '就诊时间_IN         就诊登记记录.就诊时间%TYPE,
    '状态_IN             就诊登记记录.状态%TYPE:= 0,
    '医疗类别_IN         就诊登记记录.医疗类别%TYPE:=NULL,
    '帐户余额_IN         就诊登记记录.帐户余额%TYPE:=0,
    '病种ID_IN           就诊登记记录.病种ID%TYPE:=NULL,
    '病种名称_IN         就诊登记记录.病种名称%TYPE:=NULL,
    '并发症_IN           就诊登记记录.并发症%TYPE:=NULL,
    'IC卡信息_IN         就诊登记记录.IC卡信息%TYPE:=NULL,
    'HIS流水号_IN        就诊登记记录.HIS流水号%TYPE:=NULL,
    'YB流水号_IN         就诊登记记录.YB流水号%TYPE:=NULL,
    '备注_IN             就诊登记记录.备注%TYPE:=NULL
    gstrSQL = "zl_就诊登记记录_UPDATE(" & _
        TYPE_重庆市 & "," & lng病人ID & ",0,to_date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & _
        "1," & int类别 & "," & mdbl余额 & ",NULL," & _
        "'" & str疾病编码 & "-" & str疾病名称 & "','" & str并发症 & "','" & str封锁信息 & "','" & str门诊号 & "','" & str门诊号 & "','" & mstr中心编号 & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    gstr就诊时间 = Format(datCurr, "yyyy-MM-dd HH:mm:ss")
    
    g结算数据.超限自付金额 = int类别 '用于暂时保存，门诊类别
    Call DebugTool("执行成功！")
    
    '返回格式:中间插入病人ID
    If lng病人ID <> 0 Then
        身份标识_重庆 = strIdentify & ";" & lng病人ID & str附加
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_重庆(strSelfNo As String) As Currency
'功能: 提取参保病人个人帐户余额
'参数: strSelfNO-病人个人编号
'返回: 返回个人帐户余额的金额
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHandle
    
    '从数据库中读取（因为刚才才保存了的，应该是准确的）
    If mstr医保号 = "" Or strSelfNo <> mstr医保号 Then
        gstrSQL = "Select 帐户余额 From 保险帐户 where 险类=[1] and 中心=0 and 医保号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", TYPE_重庆市, strSelfNo)
        
        If rsTemp.EOF = False Then
            个人余额_重庆 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
        End If
    Else
        个人余额_重庆 = mdbl余额
    End If
    '只能用一次
    mstr医保号 = ""
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_重庆(rs明细 As ADODB.Recordset, str结算方式 As String, Optional ByRef strAdvance As String = "") As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Static str门诊号Pre As String
    Dim str医保号 As String, StrInput As String, arrOutput  As Variant
    Dim strMessage As String
    Dim lng病人ID As Long, str规格 As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    
    Dim int单据_CUR As Integer, int单据_MAX As Integer
    Dim dbl医保基金_CUR As Double, dbl个人帐户_CUR As Double, dbl公务员补助_CUR As Double, dbl大额统筹_CUR As Double, dbl公务员返还_CUR As Double, dbl医院收支_CUR As Double, dbl民政补助_CUR As Double
    Dim dbl医保基金 As Double, dbl个人帐户 As Double, dbl公务员补助 As Double, dbl大额统筹 As Double, dbl公务员返还 As Double, dbl医院收支 As Double, dbl民政补助 As Double
    
    On Error GoTo errHandle
    
    If rs明细.RecordCount = 0 Then
        str结算方式 = "个人帐户;0;0"
        门诊虚拟结算_重庆 = True
        Exit Function
    End If
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID")
    datCurr = zlDatabase.Currentdate
    
    '分解单据总数，当前单据
    If strAdvance = "" Then strAdvance = "1|1"
    int单据_CUR = Val(Split(strAdvance, "|")(1))
    int单据_MAX = Val(Split(strAdvance, "|")(0))
    mbln多单据收费 = (int单据_MAX > 1)              '仅用来保存在保险结算记录中，用来标识是否是多单据收费
    mbln结算 = False
    Call DebugTool("多单据收费标志:" & mbln多单据收费 & ";结算标志:" & mbln结算)
    
    If mlng病人ID <> lng病人ID Then
        MsgBox "该病人还没有经过身份验证，不能进行医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mbln多单据收费 And 多单据收费_收费分别打印 Then
        MsgBox "请取消系统参数设置中票据页面里的参数“门诊收费每张单据分别打印”，方可进行多单据收费！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '如果不是多单据收费，或者多单据收费中的第一张单据，则删除已上传明细，然后清除虚拟结算保存结算结果的结构体
    '删除所有已上传的明细（指门诊业务中未结算的废记录）
    If int单据_CUR = 1 Then
        '正式结算时，strAdvance始终传的"1|1"
        mdbl费用总额 = 0
        pre_Balance.cur个人帐户 = 0
        pre_Balance.cur统筹支付 = 0
        pre_Balance.cur大额统筹 = 0
        pre_Balance.cur公务员补助 = 0
        pre_Balance.cur公务员返还 = 0
        pre_Balance.cur医院收支 = 0
        pre_Balance.cur民政补助 = 0  '20101028增加民政补助
        
        '首先退掉以前发生的所有未结的费用，包括多次执行预结算
        StrInput = "10|" & mstr门诊号 & "|" & mstr门诊号
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
        Call DebugTool("清除已上传未结算的处方明细")
    End If
    
    '保存该值
    str门诊号Pre = mstr门诊号
    
    '然后插入处方明细
    Do Until rs明细.EOF
        gstrSQL = "select A.名称,A.编码,A.类别,A.规格,A.计算单位,B.项目编码,B.附注,A.计算单位,E.规格,G.名称 剂型 " & _
                  "from 收费细目 A,保险支付项目 B,药品目录 E ,药品信息 F,药品剂型 G " & _
                  "where A.ID=[1] and A.ID=B.收费细目ID and B.险类=[2]" & _
                 "        AND A.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", CLng(rs明细("收费细目ID")), TYPE_重庆市)
        If rsTemp.EOF = True Then
            MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
            Exit Function
        End If
        mdbl费用总额 = mdbl费用总额 + Nvl(rs明细!实收金额, 0)
        If Val(Nvl(rs明细("实收金额"), 0)) <> 0 Then
            StrInput = "04|" & mstr门诊号 & "|" & mstr门诊号 & "|" & Format(datCurr, "yyyy-MM-dd HH:mm:ss")
            StrInput = StrInput & "|" & ToVarchar(rsTemp("项目编码"), 10)  '医保流水号
            StrInput = StrInput & "|" & ToVarchar(rsTemp("编码"), 20)      '医院内码
            StrInput = StrInput & "|" & ToVarchar(rsTemp("名称"), 50)      '项目名称
            StrInput = StrInput & "|" & Format(rs明细!实收金额 / Round(rs明细!数量, 2), "0.0000") '单价   可能存在打折,而记录集中单价为原始单价
            StrInput = StrInput & "|" & Format(rs明细("数量"), "0.00")     '数量
            StrInput = StrInput & "|" & IIf(rs明细("是否急诊") = 1, 1, 0)  '急诊标志
            StrInput = StrInput & "|" & Format(Nvl(rs明细!开单人, UserInfo.姓名), 20)         '处方医生
            StrInput = StrInput & "|" & Format(UserInfo.姓名, 20)          '经办人
            StrInput = StrInput & "|" & ToVarchar(rsTemp("计算单位"), 20)     '单位
            StrInput = StrInput & "|" & ToVarchar(rsTemp("规格"), 14)         '规格
            StrInput = StrInput & "|" & ToVarchar(rsTemp("剂型"), 20)         '剂理
            StrInput = StrInput & "|"                                         '冲销明细流水号
            StrInput = StrInput & "|" & Format(rs明细("实收金额"), "#####0.0000")         '金额
            
            If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
            Call AddMessage(strMessage, arrOutput, ToVarchar(rsTemp("名称"), 50), Format(rs明细!实收金额 / Round(rs明细!数量, 2), "0.0000"), False)
        End If
        rs明细.MoveNext
    Loop
    
    If strMessage <> "" Then
        strMessage = "病人费用明细传输过程中得到医保中心如下反馈信息，是否继续？" & vbCrLf & vbCrLf & strMessage
        If MsgBox(strMessage, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            '用户选择取消，先退掉明细
            StrInput = "10|" & mstr门诊号 & "|" & mstr门诊号
            If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
            
            Call DebugTool("用户选择取消，退掉明细")
            Exit Function
        End If
    End If
    
    '调用预结算
    StrInput = "06|" & mstr门诊号 & "|||" & Format(mdbl费用总额, "#0.00;-#0.00;0;")
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    '赋值
    dbl个人帐户 = Val(arrOutput(2))
    dbl医保基金 = Val(arrOutput(1))
    dbl公务员补助 = Val(arrOutput(3))
    dbl大额统筹 = Val(arrOutput(5))
    dbl公务员返还 = Val(arrOutput(6))
    If UBound(arrOutput) > 7 Then
        dbl医院收支 = Val(arrOutput(8))
    End If
    If UBound(arrOutput) > 8 Then   '20101028增加民政补助
        dbl民政补助 = Val(arrOutput(9))
        dbl医保基金 = dbl医保基金 - dbl民政补助
    End If
    
    Call DebugTool("获取本次预结算结果")
    
    '计算本次真实的结算结果（处理多单据收费的情况）
    dbl个人帐户_CUR = dbl个人帐户 - pre_Balance.cur个人帐户
    dbl医保基金_CUR = dbl医保基金 - pre_Balance.cur统筹支付
    dbl大额统筹_CUR = dbl大额统筹 - pre_Balance.cur大额统筹
    dbl公务员补助_CUR = dbl公务员补助 - pre_Balance.cur公务员补助
    dbl公务员返还_CUR = dbl公务员返还 - pre_Balance.cur公务员返还
    dbl医院收支_CUR = dbl医院收支 - pre_Balance.cur医院收支
    dbl民政补助_CUR = dbl民政补助 - pre_Balance.cur民政补助  '20101028增加民政补助
    Call DebugTool("得到本次预结算的真实结果")
    
    '返回结算结果（多单据返回差额，单张收费本次的结果就是差额）
    str结算方式 = "个人帐户;" & dbl个人帐户_CUR & ";0"   '不能修改个人帐户，因为结算时已经不再传金额到前置机了
    If Val(Format(dbl医保基金_CUR, "#0.00")) <> 0 Then
        str结算方式 = str结算方式 & "|医保基金;" & dbl医保基金_CUR & ";0"
    End If
    If Val(Format(dbl公务员补助_CUR, "#0.00")) <> 0 Then
        str结算方式 = str结算方式 & "|公务员补助;" & dbl公务员补助_CUR & ";0"
    End If
    If Val(Format(dbl大额统筹_CUR, "#0.00")) <> 0 Then
        str结算方式 = str结算方式 & "|大额统筹;" & dbl大额统筹_CUR & ";0"
    End If
    If Val(Format(dbl公务员返还_CUR, "#0.00")) <> 0 Then
        str结算方式 = str结算方式 & "|公务员返还;" & dbl公务员返还_CUR & ";0"
    End If
    If Val(Format(dbl民政补助_CUR, "#0.00")) <> 0 Then    '20101028增加民政补助
        str结算方式 = str结算方式 & "|民政补助;" & dbl民政补助_CUR & ";0"
    End If
    If UBound(arrOutput) > 7 Then
        str结算方式 = str结算方式 & "|医院收支;" & dbl医院收支_CUR & ";0"
    End If
    Call DebugTool("返回结算结果串:" & str结算方式)
    
    '保存累计值
    pre_Balance.cur个人帐户 = dbl个人帐户
    pre_Balance.cur统筹支付 = dbl医保基金
    pre_Balance.cur大额统筹 = dbl大额统筹
    pre_Balance.cur公务员补助 = dbl公务员补助
    pre_Balance.cur公务员返还 = dbl公务员返还
    pre_Balance.cur医院收支 = dbl医院收支
    pre_Balance.cur民政补助 = dbl民政补助   '20101028增加民政补助
    Call DebugTool("保存累计值")
    
    门诊虚拟结算_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_重庆(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, Optional strAdvance As String, Optional bln门诊 As Boolean = True) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim str医保号 As String, StrInput As String
    Dim lng病人ID  As Long
    Dim str操作员 As String, arrOutput  As Variant
    Dim datCurr As Date
    Dim str病种 As String, str并发症 As String, str封锁信息 As String, str备注 As String
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Dim bln帐户累计 As Boolean
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    
    Dim cur统筹支付 As Double
    Dim cur公务员补助 As Double
    Dim cur大额统筹 As Double
    Dim cur医院收支 As Double
    Dim cur公务员返还 As Double, cur民政补助 As Double, cur民政帐户余额 As Double '20101028增加民政补助
    Dim cur发生费用 As Currency
    
    Dim str结算方式 As String
    Dim blnOld As Boolean
    Dim blnRevise As Boolean
    Dim bln帐户支付 As Boolean
    Dim int帐户支付方式 As Integer '0-支付;1-住院询问;2-门诊询问;3-不支付
    
    On Error GoTo errHandle
    
    Call DebugTool("进入门诊结算")
    strAdvance = ""
    gstrSQL = "Select * From 门诊费用记录 Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", lng结帐ID)
    If rs明细.EOF = True Then
        Err.Raise 9000 + VbMsgBoxStyle.vbExclamation, gstrSysName, "没有填写收费记录"
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str操作员 = ToVarchar(IIf(IsNull(rs明细("操作员姓名")), UserInfo.姓名, rs明细("操作员姓名")), 20)
    
    If mlng病人ID <> lng病人ID Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "该病人还没有经过身份验证，不能进行医保结算。"
        Exit Function
    End If
    
''   不必统计费用总额了，以预结算累计的为准
'    Do Until rs明细.EOF
'        cur发生费用 = cur发生费用 + rs明细("结帐金额")
'        Call TrackRecordInsure(rs明细!ID, rs明细!收费细目ID)
'        rs明细.MoveNext
'    Loop
'    cur发生费用 = Val(Format(cur发生费用, "#0.00;-#0.00;0;"))
    
    '读取参保病人的并发症及封锁信息
    gstrSQL = "Select 退休证号 As 病种编码,病种名称,并发症,封锁信息" & _
        " From 保险帐户 " & _
        " Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取参保病人的并发症及封锁信息", lng病人ID, TYPE_重庆市)
    If Not rsTemp.EOF Then
        str病种 = Nvl(rsTemp!病种编码)
        If str病种 <> "" Then str病种 = "[" & str病种 & "]"
        str病种 = str病种 & Nvl(rsTemp!病种名称)
        str并发症 = Nvl(rsTemp!并发症)
        str封锁信息 = Nvl(rsTemp!封锁信息)
    End If
    str备注 = str病种 & "||" & str并发症 & "||" & str封锁信息 & "@@本次是" & IIf(mbln多单据收费, "多单据收费", "单张收费")
    Call DebugTool("成功获取病人相关信息——病种编码与名称、并发症及封锁信息")
    
    '调用结算
    bln帐户支付 = True
    If mbln结算 = False Then
        '如果不是多单据收费，则询问是否进行个人帐户支付
        If Not mbln多单据收费 Then
            '取保险参数“离休就诊录入诊断”
            gstrSQL = "Select Nvl(参数值,0) AS 参数值 From 保险参数 Where 险类=[1] And 参数名='个人帐户'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险参数“离休就诊录入诊断”", TYPE_重庆市)
            If rsTemp.RecordCount <> 0 Then
                int帐户支付方式 = rsTemp!参数值
            End If
            If int帐户支付方式 < 2 Then
                bln帐户支付 = True
            ElseIf int帐户支付方式 = 2 Then
                If MsgBox("请问本次要进行个人帐户支付吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    bln帐户支付 = False
                Else
                    bln帐户支付 = True
                End If
            Else
                bln帐户支付 = False
            End If
        End If
        
        Call DebugTool("准备调用门诊结算")
        StrInput = "05|" & mstr门诊号 & "|1||" & str操作员 & "|" & IIf(bln帐户支付, "0", "1") & "||" & mdbl费用总额 '用帐户余额支付
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
        mbln结算 = True
        bln帐户累计 = True
        
        '保存结算记录
        '---------------------------------------------------------------------------------------------
        mstr结算流水号 = arrOutput(1)
        cur统筹支付 = Val(arrOutput(2))
        cur个人帐户 = Val(arrOutput(3))
        cur公务员补助 = Val(arrOutput(4))
        cur公务员返还 = Val(arrOutput(7))
        cur大额统筹 = Val(arrOutput(6))
        If UBound(arrOutput) > 8 Then
            cur医院收支 = Val(arrOutput(9))
        End If
        If UBound(arrOutput) > 9 Then   '20101028增加民政补助
            cur民政补助 = Val(arrOutput(10))
            cur统筹支付 = cur统筹支付 - cur民政补助
            cur民政帐户余额 = Val(arrOutput(11))
            gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'民政救助门诊余额','''" & cur民政帐户余额 & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新民政救助门诊余额")
        End If
        
        
        Call DebugTool("得到结算结果")
        
        '保险结算记录的备注中保存本次是否多单据收费，以及是否是第一张单据
        str备注 = str备注 & "||1"
    
        '更新就诊登记记录的记录ID，以便与保险结算记录关联
        gstrSQL = "zl_就诊登记记录_结束(" & TYPE_重庆市 & "," & lng病人ID & ",0," & _
            "to_date('" & gstr就诊时间 & "','yyyy-MM-dd hh24:mi:ss')," & lng结帐ID & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        '如果不是多单据收费
        If Not mbln多单据收费 And bln门诊 Then
            '只会出现个人帐户实际支付与虚拟结算不符的情况，原因如下：
            '1、如果预结算后长时间不结算，有可能中心将其帐户进行了特殊保护处理，不能再下帐
            '2、第一次结算失败，但帐户已下，再次点确定时由于帐户余额为零导致第二次下卡为零，从而出现东软虚拟结算与实际结算不符的情况
            If pre_Balance.cur个人帐户 <> cur个人帐户 Or pre_Balance.cur统筹支付 <> cur统筹支付 Then
                blnRevise = True
                
                str结算方式 = "个人帐户|" & cur个人帐户
                str结算方式 = str结算方式 & "||医保基金|" & cur统筹支付
                If cur大额统筹 <> 0 Then str结算方式 = str结算方式 & "||大额统筹|" & cur大额统筹
                If cur公务员补助 <> 0 Then str结算方式 = str结算方式 & "||公务员补助|" & cur公务员补助
                If cur公务员返还 <> 0 Then str结算方式 = str结算方式 & "||公务员返还|" & cur公务员返还
                If cur民政补助 <> 0 Then str结算方式 = str结算方式 & "||民政补助|" & cur民政补助    '20101028增加民政补助
                If cur医院收支 <> 0 Then str结算方式 = str结算方式 & "||医院收支|" & cur医院收支
                If str结算方式 <> "" Then
                    #If gverControl < 2 Then
                        blnOld = True
                        gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
                    #Else
                        strAdvance = str结算方式
                        gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
                    #End If
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
                End If
            End If
        End If
    Else
        '第一张单据保存所有的总额，其后的多张单据，费用总额及医保金额保存为零
        If mbln多单据收费 = False Then
            MsgBox "警告：请操作员截图、记录本次操作流程，并与重庆中联公司服务部联系！", vbInformation, gstrSysName
        End If
        
        mdbl费用总额 = 0
        cur统筹支付 = 0
        cur个人帐户 = 0
    End If
    
    '帐户年度信息
    datCurr = zlDatabase.Currentdate
    
    If bln帐户累计 Then
        Call DebugTool("保存帐户年度信息")
        Call Get帐户信息(TYPE_重庆市, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_重庆市 & "," & Year(datCurr) & "," & _
            cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
            cur进入统筹累计 + cur统筹支付 & "," & _
            cur统筹报销累计 + cur统筹支付 & "," & int住院次数累计 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    End If
    
    'g结算数据.超限自付金额中保存的是门诊病人就诊类型（急诊、特殊病门诊或普通门诊），结算记录的备注保存的是病种的名称
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)'超限自付金额用于暂时保存，门诊类别
    Call DebugTool("保存保险结算记录")
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_重庆市 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & mdbl费用总额 & ",0,0," & _
        cur统筹支付 & "," & cur统筹支付 & "," & cur民政补助 & "," & g结算数据.超限自付金额 & "," & cur个人帐户 & ",'" & mstr结算流水号 & "',NULL,NULL,'" & str备注 & "'" & _
        IIf(blnOld, "", IIf(blnRevise, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    门诊结算_重庆 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_重庆(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur票据总金额 As Currency
    Dim curDate As Date
        
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    '此代码不能注释，要取病人ID，不然其后更新帐户年度信息时要报错
    gstrSQL = "Select * From 门诊费用记录 " & _
        " Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", lng结帐ID)
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")

        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    cur票据总金额 = Val(Format(cur票据总金额, "#####0.00"))
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")
'
'    '保存医保项目的状态，备查
'    gstrSQL = "Select * From 病人费用记录 " & _
'        " Where 结帐ID=" & lng冲销ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
'    Call OpenRecordset(rsTemp, "重庆医保")
'    Do While Not rsTemp.EOF
'        Call TrackRecordInsure(rsTemp!ID, rsTemp!收费细目ID)
'        rsTemp.MoveNext
'    Loop
    
    Call 多单据收费_退费(lng结帐ID)
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", TYPE_重庆市, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    str流水号 = rsTemp("支付顺序号")
    cur个人帐户 = Nvl(rsTemp!个人帐户支付, 0)
    
    '如果是多单据收费，且不是第一张单据，则直接返回真
    If InStr(rsTemp!备注, "@@") <> 0 Then
        If UBound(Split(Split(rsTemp!备注, "@@")(1), "||")) > 0 Then
            If Val(Split(Split(rsTemp!备注, "@@")(1), "||")(1)) = 0 Then
                门诊结算冲销_重庆 = True
            End If
        Else
            门诊结算冲销_重庆 = True
        End If
    End If
    
    If 门诊结算冲销_重庆 = False Then
        StrInput = "99|" & str流水号 & "|" & ToVarchar(UserInfo.姓名, 20)
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_重庆市, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_重庆市 & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_重庆市 & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & rsTemp!发生费用金额 * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") & "," & _
        cur个人帐户 * -1 & ",'" & str流水号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")

    门诊结算冲销_重庆 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 个人帐户转预交_重庆(lng预交ID As Long, cur个人帐户 As Currency, strSelfNo As String, str顺序号 As String, ByVal lng病人ID As Long) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    
    个人帐户转预交_重庆 = False
End Function

Public Function 入院登记_重庆(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim StrInput As String, arrOutput  As Variant
    Dim datCurr As Date, rsTemp As New ADODB.Recordset
    Dim str卡号 As String, str顺序号 As String
    Dim strTemp As String, str提示 As String, str诊断 As String
    Dim str住院门诊号 As String
    Dim str医疗类别 As String
    
    On Error GoTo errHandle
    
    '获得病人出院诊断
    gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=[1] and A.主页ID=[2]" & _
              " and A.诊断类型=1 and A.诊断次序=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "入院登记", lng病人ID, lng主页ID)
    If rsTemp.EOF = False Then
        str诊断 = ToVarchar(rsTemp("描述信息"), 40)
    End If
    
    '获得医保号
    gstrSQL = "select 卡号,医保号 from 保险帐户 where 险类=[1] and 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "入院登记", TYPE_重庆市, lng病人ID)
    str卡号 = IIf(IsNull(rsTemp("卡号")), "", rsTemp("卡号"))
    str医保号 = rsTemp("医保号")
    
    '不是以下情况（不存在字段“住院门诊号”，或者值为空）则强制将住院门诊号清为空
    str住院门诊号 = ""
    If GetMode(lng病人ID, lng主页ID, str住院门诊号) = False Then
        gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'住院门诊号','NULL')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
    End If
    
    '获得其它入院信息
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.入院方式,nvl(A.二级院转入,0) as 二级院转入,A.门诊医师,A.入院日期,A.入院病床,B.名称 as 入院科室 from 病案主页 A,部门表 B " & _
             " Where A.入院科室ID = B.ID And A.病人ID =[1] And A.主页ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "入院登记", lng病人ID, lng主页ID)
    '取医疗类别
    str医疗类别 = IIf(rsTemp!入院方式 = "家床", "23", IIf(rsTemp!入院方式 = "转入", 22, 21))
'    If Is家庭病床(lng病人ID, lng主页ID) Then
'        str医疗类别 = "23"
'    Else
'        str医疗类别 = IIf(rsTemp!入院方式 = "转入", 22, 21)
'    End If
    
    '调用入院接口
    StrInput = "02|" & GetIdentify(lng病人ID, lng主页ID, True) & "|" & str医疗类别 & "|" & str医保号 & "|" & _
               ToVarchar(rsTemp("入院科室"), 30) & "|" & ToVarchar(rsTemp("门诊医师"), 20) & "|" & _
               Format(rsTemp("入院日期"), "yyyy-MM-dd") & "|" & ToVarchar(str诊断, 40) & "|" & ToVarchar(UserInfo.姓名, 20) & "|0"
    Call DebugTool("调东软入院接口：" & StrInput)
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    str顺序号 = arrOutput(1)
    mdbl余额 = Val(arrOutput(2))
    Call DebugTool("东软返回流水号与帐户余额：" & str顺序号 & "；" & mdbl余额)
    
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_重庆市 & "," & Year(datCurr) & "," & _
        mdbl余额 & ",0,0,0,0,0,0,0,0,0,'')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    '强制把登记顺序号、及新的医保号填入
    gstrSQL = "ZL_保险帐户_修改医保号(" & lng病人ID & "," & TYPE_重庆市 & _
                ",'" & str卡号 & "','" & str医保号 & "','" & str顺序号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    入院登记_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_重庆(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
            '取入院登记验证所返回的顺序号
    Dim datCurr As Date, rsTemp As New ADODB.Recordset
    Dim StrInput As String, arrOutput  As Variant, bln零费用出院 As Boolean
    Dim str诊断 As String
    Dim str住院门诊号 As String
    
    On Error GoTo errHandle
    
    '获得病人出院诊断
    gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=[1] and A.主页ID=[2]" & _
              " and A.诊断类型=3 and A.诊断次序=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院登记", lng病人ID, lng主页ID)
    If rsTemp.EOF = False Then
        str诊断 = Nvl(rsTemp("描述信息"), "无")
    Else
        str诊断 = "无"   '诊断不论如何不能为空
    End If
    str诊断 = ToVarchar(str诊断, 40)
    
    '获得其它出院信息
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.住院医师,A.入院日期,A.出院日期,A.出院病床,B.名称 as 出院科室 from 病案主页 A,部门表 B " & _
             " Where A.出院科室ID = B.ID And A.病人ID =[1] And A.主页ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院登记", lng病人ID, lng主页ID)
    '调用接口，更新其住院信息
    StrInput = "03|" & GetIdentify(lng病人ID, lng主页ID) & "|0000010010|21|||" & Format(rsTemp("入院日期"), "yyyy-MM-dd") & "||" & _
                Format(rsTemp("出院日期"), "yyyy-MM-dd") & "|||" & ToVarchar(UserInfo.姓名, 20) & "|0"
    
    '检查该次住院是否没有费用发生
    gstrSQL = "Select nvl(sum(实收金额),0) as 金额  from 住院费用记录 where 病人ID=[1] and 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病人出院", lng病人ID, lng主页ID)
    If rsTemp.EOF = True Then
        bln零费用出院 = True
    Else
        bln零费用出院 = (rsTemp("金额") = 0)
    End If
    
    If bln零费用出院 = True Then
        '对于零费用出院，就将其处理为退入院。而不用更新其住院信息
        gstrSQL = "Select 顺序号 from 保险帐户 where 病人ID=[1] and 险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病人出院", lng病人ID, TYPE_重庆市)
        StrInput = "99|" & rsTemp("顺序号") & "|" & ToVarchar(UserInfo.姓名, 20)
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    End If
    
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重庆市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    出院登记_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_重庆(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '检查是否存在未结费用，不存在未结费用的病人不允许撤销出院
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        MsgBox "该病人不存在未结费用，不允许办理撤销出院！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If MsgBox("你确定要将该病人撤销出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    出院登记撤销_重庆 = True
End Function

Public Function 转高龄家庭病床(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '调用接口更新：住院号(门诊号)|更新标志|医疗类别|科室|医生|入院日期|入院诊断|出院日期|确诊疾病编码
                  '|出院原因|经办人|并发症
    Dim StrInput As String
    Dim arrOutput
    
    StrInput = "03|" & GetIdentify(lng病人ID, lng主页ID) & "|1000000000|24|||||||||"
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'高龄病床','''" & "1" & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "高龄病床标志更新为1")
End Function

Public Function 更新出院疾病_重庆(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：更新病人的出院疾病。如果是肿瘤，则结算时起付线会减半
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str分类 As String, str并发症 As String, str疾病编码 As String, str疾病名称 As String
    Dim StrInput As String, arrOutput  As Variant
    Dim str科室 As String, str医生 As String, bln高龄病床 As Boolean
    Dim str入院日期 As String, str出院日期 As String, str诊断 As String, str医疗类别 As String
    
    On Error GoTo errHandle
    
    '获得病人出院病种及并发症
    gstrSQL = "Select 退休证号 病种编码,病种名称,并发症,高龄病床 From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人出院病种及并发症", lng病人ID)
    str疾病编码 = Nvl(rsTemp!病种编码)
    str并发症 = Nvl(rsTemp!并发症)
    str疾病名称 = Nvl(rsTemp!病种名称)
    bln高龄病床 = (Nvl(rsTemp!高龄病床, "0") = "1")
    
    str分类 = "出院"
    If frm疾病选择重庆.GetCode("", str分类, str疾病编码, str疾病名称, str并发症) = False Then
        Exit Function
    End If
    str疾病编码 = ToVarchar(str疾病编码, 20)
    str并发症 = ToVarchar(str并发症, 200)
    str疾病名称 = TrimStr(str疾病名称)
    
    '获得病人出院诊断
    gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=[1] and A.主页ID=[2]" & _
              " and A.诊断类型=1 and A.诊断次序=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "入院登记", lng病人ID, lng主页ID)
    If rsTemp.EOF = False Then
        str诊断 = ToVarchar(rsTemp("描述信息"), 40)
    End If
    
    '取病人的入院信息
    gstrSQL = "select A.入院方式,nvl(A.二级院转入,0) as 二级院转入,A.门诊医师,A.入院日期,A.出院日期,A.入院病床,B.名称 as 入院科室 from 病案主页 A,部门表 B " & _
             " Where A.入院科室ID = B.ID And A.病人ID = [1] And A.主页ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人的入院信息", lng病人ID, lng主页ID)
    If bln高龄病床 = False Then
        str医疗类别 = IIf(rsTemp!入院方式 = "家床", "23", IIf(rsTemp!入院方式 = "转入", 22, 21))
'        If Is家庭病床(lng病人ID, lng主页ID) Then
'            str医疗类别 = "23"
'        Else
'            str医疗类别 = IIf(rsTemp!入院方式 = "转入", 22, 21)
'        End If
    Else
        str医疗类别 = "24"
    End If
    str科室 = ToVarchar(rsTemp("入院科室"), 30)
    str医生 = ToVarchar(rsTemp("门诊医师"), 20)
    str入院日期 = Format(rsTemp!入院日期, "yyyy-MM-dd")
    If IsNull(rsTemp!出院日期) Then
        str出院日期 = ""
    Else
        str出院日期 = Format(rsTemp!出院日期, "yyyy-MM-dd")
    End If
    
    'Modified by ZYB 2004-05-10
    '调用接口更新：住院号(门诊号)|更新标志|医疗类别|科室|医生|入院日期|入院诊断|出院日期|确诊疾病编码
                  '|出院原因|经办人|并发症
    StrInput = "03|" & GetIdentify(lng病人ID, lng主页ID) & "|1110111001|" & str医疗类别 & "|" & str科室 & _
               "|" & str医生 & "|" & str入院日期 & "|" & str诊断 & "|" & str出院日期 & "|" & str疾病编码 & "|||" & str并发症
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'退休证号','''" & str疾病编码 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'病种名称','''" & str疾病名称 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'并发症','''" & str并发症 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新并发症")
    
    更新出院疾病_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 撤消医保入院_重庆(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str顺序号 As String) As Boolean
'功能：更新病人的出院疾病。如果是肿瘤，则结算时起付线会减半
    Dim StrInput As String, arrOutput  As Variant
    
    On Error GoTo errHandle
    
    '调用接口
    StrInput = "99|" & str顺序号 & "|" & ToVarchar(UserInfo.姓名, 20)
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    gstrSQL = "ZL_病案主页_撤消医保入院(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "撤消医保入院")
    
    撤消医保入院_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_重庆(rsExse As Recordset, ByVal lng病人ID As Long, ByVal str医保号 As String, _
        Optional ByVal bln费用查询 As Boolean = False, Optional ByRef strAdvance As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim cn上传 As New ADODB.Connection, rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    
    Dim str封锁信息 As String
    Dim StrInput As String, arrOutput   As Variant, arrTemp As Variant
    Dim cur发生费用 As Double
    Dim str总金额医院 As String, str总金额医保 As String, str冲销明细流水号 As String
    Dim str医生 As String, datCurr As Date, intMsg As Integer
    Dim bln单病种 As Boolean, bln中途结算只处理已上传费用 As Boolean
    Dim bln批量预结算 As Boolean
    
    Dim str入院日期 As String, str上次结算日期 As String
    
    On Error GoTo errHandle
    If strAdvance = "" Then strAdvance = "0"
    bln批量预结算 = (Val(Split(strAdvance, ";")(0)) = 1)
    strAdvance = ""         '对此参数的赋值，将在批量预结算完成后显示给前台，所以此处先清为空
    
    mlng病人ID = 0         '初始化。只要一选择病人，就会调用本过程，也就会设成0
    mstr费用终止日期 = ""
    pre_Balance.HIS计算医院收支 = False

    If rsExse.RecordCount = 0 Then
        strAdvance = MessageInfo("该病人没有有发生费用，无法进行结算操作。", bln批量预结算)
        Exit Function
    End If
    rsExse.MoveFirst
    
    datCurr = zlDatabase.Currentdate
    With g结算数据
        .病人ID = rsExse("病人ID")
        
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", CLng(rsExse("病人ID")))
        If IsNull(rsTemp("主页ID")) = True Then
            strAdvance = MessageInfo("只有住院病人才可以使用医保结算。", bln批量预结算)
            Exit Function
        End If
        .主页ID = rsTemp("主页ID")
        .年度 = Int(Format(datCurr, "yyyy"))
        
        '判断中途结算是否只结算已上传费用
        bln中途结算只处理已上传费用 = True
        gstrSQL = " Select Nvl(参数值,1) AS 参数值 From 保险参数 Where 险类=[1] And 参数名='中途结算'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断中途结算是否只结算已上传费用", TYPE_重庆市)
        If rsTemp.RecordCount <> 0 Then
            bln中途结算只处理已上传费用 = (rsTemp!参数值 = 1)
        End If
    End With
    
    'Modified by ZYB 2004-05-10
    '提取病人的基本信息，如果存在封锁原因则提示
    gstrSQL = "Select 医保号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该病人的医保号", TYPE_重庆市, g结算数据.病人ID)
    
    StrInput = "01|" & rsTemp!医保号
    If HandleBusiness(StrInput, arrTemp, bln批量预结算) = False Then Exit Function
    str封锁信息 = ""
    If Val(arrTemp(11)) <> 0 Then
        str封锁信息 = arrTemp(12)
        If Not bln批量预结算 Then MsgBox str封锁信息, vbInformation, gstrSysName
    End If
    '更新封锁信息
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'封锁信息','''" & str封锁信息 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
    
    Screen.MousePointer = vbHourglass
    '1.2 读出病人的入院时间
    ''费用终止日期：如果已出院，即是出院日期，否则是本次结帐中最大的发生日期
    gstrSQL = "select 入院日期,nvl(出院日期,to_date('3000-01-01','yyyy-MM-dd')) as 出院日期 " & _
              "from 病案主页 where 病人ID=[1] and 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", g结算数据.病人ID, g结算数据.主页ID)
    str入院日期 = Format(rsTemp!入院日期, "yyyy-MM-dd")
    If rsTemp("出院日期") = CDate("3000-01-01") Then
        g结算数据.中途结帐 = 1
        With rsExse
            Do While Not .EOF
                If Format(!发生时间, "yyyy-MM-dd") > mstr费用终止日期 Then mstr费用终止日期 = Format(!发生时间, "yyyy-MM-dd")
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
        End With
    Else
        g结算数据.中途结帐 = 0
        mstr费用终止日期 = Format(rsTemp("出院日期"), "yyyy-MM-dd")
    End If
    
    '1.3 读出本次住院上次中结时间，如果为空，上次中结结束时间为入院日期
     gstrSQL = " Select Max(发生时间) AS 结束日期 From 住院费用记录" & _
               " Where 结帐Id=(" & _
               "    Select Max(ID) As 结帐ID" & _
               "    From 病人结帐记录" & _
               "    Where 病人ID=[1] And 记录状态 = 1 " & _
               "    And 收费时间>[2]" & _
               "    And Nvl(附加标志,0)<>9"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读出本次住院上次中结时间", g结算数据.病人ID, CDate(str入院日期))
    If rsTemp.RecordCount = 0 Then
        str上次结算日期 = str入院日期
    Else
        If IsNull(rsTemp!结束日期) Then
            str上次结算日期 = str入院日期
        Else
            str上次结算日期 = Format(rsTemp!结束日期, "yyyy-MM-dd")
        End If
    End If
    
    '1.4 中途结算算进不算出，直接是两个结帐日期相减；出院结算需加1天
    '本次结帐费用中最大的发生日期减去上次结帐费用中最大的发生日期，即是住院床日
    g结算数据.住院床日 = DateDiff("d", CDate(str上次结算日期), CDate(mstr费用终止日期))
    'If g结算数据.中途结帐 = 0 Then g结算数据.住院床日 = g结算数据.住院床日 + 1
    If g结算数据.住院床日 < 1 Then g结算数据.住院床日 = 1 '至少有一天
    Call DebugTool("费用终止日期：" & mstr费用终止日期 & "；上次结算日期或入院日期：" & str上次结算日期 & "；住院床日：" & g结算数据.住院床日)
    
    Do Until rsExse.EOF
        cur发生费用 = cur发生费用 + rsExse("金额")
        rsExse.MoveNext
    Loop
    cur发生费用 = Val(Format(cur发生费用, "#####0.00"))
    
    '只有出院结算或使用病人费用查询处调用，才上传所有未上传明细，中途结算只对已上传数据进行结算
    If g结算数据.中途结帐 = 0 Or bln费用查询 Or Not bln中途结算只处理已上传费用 Then
        '读出未上传明细（排序，以便先上传正明细，再上传负明细）
        gstrSQL = "Select A.ID,A.NO,A.记录性质,A.记录状态,A.序号,A.病人ID,A.主页ID,A.发生时间 as 登记时间,Round(A.实收金额,4) 实收金额" & _
                  "         ,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
                  "         ,C.项目编码,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,E.规格,G.名称 剂型 " & _
                  "  From 住院费用记录 A,收费细目 B,保险支付项目 C,病案主页 D,药品目录 E ,药品信息 F,药品剂型 G " & _
                  "  where A.病人ID=" & lng病人ID & " and A.主页ID=" & g结算数据.主页ID & _
                  "        and A.记帐费用=1 and A.记录性质<>1 and Nvl(A.实收金额,0)<>0 and nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 " & _
                  "        and A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=[1]" & _
                  "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类=D.险类 " & _
                  "        AND B.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) " & _
                  "  Order by A.发生时间,A.记录性质,Decode(A.记录状态,2,2,1)"
        Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", TYPE_重庆市)
'
'        '保存医保项目的信息，备查
'        Do While Not rs明细.EOF
'            Call TrackRecordInsure(rs明细!ID, rs明细!收费细目ID)
'            rs明细.MoveNext
'        Loop
        
        '打开另外一个连接串，以达到不受当前连接事务的控制
        Set cn上传 = GetNewConnection
        cn上传.Open
        
        intMsg = 0
        Call DebugTool("开始上传明细")
        If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
        Do Until rs明细.EOF
            If rs明细!记录状态 = 1 Then
                If Val(rs明细!数量) < 0 Or Val(rs明细!价格) < 0 Then
                    '任意取一笔正常记录的流水号，做为冲销流水号
                    '改为标志，该变量不为空
                    str冲销明细流水号 = "冲销记录"
                Else
                    str冲销明细流水号 = ""
                End If
            Else
                str冲销明细流水号 = GetDetailSequence(rs明细!NO, rs明细!序号, rs明细!记录性质, rs明细!记录状态)
            End If
            
            If rs明细!记录状态 = 1 And str冲销明细流水号 <> "" Then
                Call UploadNegative(rs明细!病人ID, rs明细!主页ID, rs明细!ID, rs明细!收费细目ID)
            Else
                If rs明细!记录状态 <> 2 Then
                    str医生 = ToVarchar(IIf(IsNull(rs明细("医生")), UserInfo.姓名, rs明细("医生")), 20)
                    
                    StrInput = "04|" & GetIdentify(lng病人ID, g结算数据.主页ID)
                    StrInput = StrInput & "|" & rs明细("NO") & "_" & rs明细("记录性质")
                    StrInput = StrInput & "|" & Format(rs明细("登记时间"), "yyyy-MM-dd HH:mm:ss")
                    StrInput = StrInput & "|" & ToVarchar(rs明细("项目编码"), 10) '中心编码
                    StrInput = StrInput & "|" & ToVarchar(rs明细("编码"), 20) '医院内码
                    StrInput = StrInput & "|" & ToVarchar(rs明细("名称"), 50)     '项目名称
                    StrInput = StrInput & "|" & Format(rs明细("价格"), "0.0000")      '单价
                    StrInput = StrInput & "|" & Format(rs明细("数量"), "0.00")        '数量
                    StrInput = StrInput & "|" & IIf(rs明细("是否急诊") = 1, 1, 0)     '急诊标志
                    StrInput = StrInput & "|" & str医生                               '医生
                    StrInput = StrInput & "|" & ToVarchar(UserInfo.姓名, 20)          '经办人
                    StrInput = StrInput & "|" & ToVarchar(rs明细("计算单位"), 20)     '单位
                    StrInput = StrInput & "|" & ToVarchar(rs明细("规格"), 14)         '规格
                    StrInput = StrInput & "|" & ToVarchar(rs明细("剂型"), 20)         '剂理
                    StrInput = StrInput & "|" & str冲销明细流水号                     '冲销明细流水号
                    StrInput = StrInput & "|" & Format(rs明细("实收金额"), "#####0.0000")         '金额
                Else
                    '记帐表和记帐单都允许部分冲销后，需要一笔一笔冲销
                    StrInput = "99|" & str冲销明细流水号 & "|" & ToVarchar(UserInfo.姓名, 20)
                End If
                
                'Modified by ZYB 20040511 昆明处理
                '对于由于负数记帐，产生的冲销记录，因为负数记帐那笔由于接口限制，肯定传不上去，因此本次限制它产生的冲销记录不上传
                If HandleBusiness(StrInput, arrOutput, bln批量预结算) = False Then
                    '费用上传失败
                    If Not bln批量预结算 Then
                        If MsgBox("[病人ID:" & lng病人ID & "]单据“" & rs明细("NO") & "”中" & rs明细("名称") & "费用上传失败，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                        If intMsg = 0 Then
                            If MsgBox("上传数据失败，是否停止数据上传并直接进行结帐？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                                intMsg = 1
                                Exit Do
                            Else
                                intMsg = -1
                            End If
                        End If
                    Else
                        strAdvance = MessageInfo("[病人ID:" & lng病人ID & "]单据“" & rs明细("NO") & "”中" & rs明细("名称") & "费用上传失败", bln批量预结算)
                        Exit Function
                    End If
                Else
                    '费用上传成功才做上标记
                    If rs明细!记录状态 <> 2 Then
                        gstrSQL = "ZL_病人记帐记录_上传(" & rs明细("ID") & "," & Val(arrOutput(2)) * rs明细("数量") & ",'" & arrOutput(1) & "')"
                    Else
                        gstrSQL = "ZL_病人记帐记录_上传(" & rs明细("ID") & ")"
                    End If
                    cn上传.Execute gstrSQL, , adCmdStoredProc
                End If
            End If
            
            rs明细.MoveNext
        Loop
    End If
    
    '如果是出院结算前的虚拟结算，检查是否存在未上传的费用明细
    If g结算数据.中途结帐 = 0 Then
        If Not frm查询未上传费用明细.ShowME(lng病人ID, g结算数据.主页ID, TYPE_重庆市) Then Exit Function
    End If
    
    '调用预结算
    Call DebugTool("准备调用预结算")
    '病人在进行出院结算前的虚拟结算时，费用终止日期传当前日期；如果是进行中途结算的虚拟结算则传费用记录中最大的登记日期，有利于核对两边的费用是否一致
    '原因，出院结算时传递的费用终止日期无效，东软仍然统计的所有费用进行的结算
    If g结算数据.中途结帐 = 0 Then
        mstr费用终止日期 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    End If
    
    '住院(门诊)号|截止日期|住院床日，与周韬商量，以登记时间为准，取最大的登记时间做为截止日期
    StrInput = "06|" & GetIdentify(lng病人ID, g结算数据.主页ID) & "|" & Format(mstr费用终止日期, "yyyyMMdd") & "|" & g结算数据.住院床日 & "|" & Format(cur发生费用, "#0.00")
    If HandleBusiness(StrInput, arrOutput, bln批量预结算) = False Then Exit Function
    
    pre_Balance.cur个人帐户 = Val(arrOutput(2))
    pre_Balance.cur统筹支付 = Val(arrOutput(1))
    pre_Balance.cur大额统筹 = Val(arrOutput(5))
    pre_Balance.cur公务员补助 = Val(arrOutput(3))
    pre_Balance.cur公务员返还 = Val(arrOutput(6))
    pre_Balance.cur医院收支 = 0
    '在新接口增加医院收支时保存其值，临时用于总额比较
    If UBound(arrOutput) > 7 Then
        pre_Balance.cur医院收支 = Val(arrOutput(8))
    End If
    If UBound(arrOutput) > 8 Then       '20101028增加民政补助
        pre_Balance.cur民政补助 = Val(arrOutput(9))
        pre_Balance.cur统筹支付 = pre_Balance.cur统筹支付 - pre_Balance.cur民政补助
    End If
    
    '保存病人个人帐户余额
    mstr医保号 = str医保号
    mdbl余额 = Val(arrOutput(7)) + Val(arrOutput(2))
    
    '保存临时数据，为结算操作做准备
    With g结算数据
        .发生费用金额 = cur发生费用
    End With
    
    str总金额医院 = Format(cur发生费用, "#####0.00")
    '东软返回的现金支付应该减去公务员返还部分，才是最终的现金支付额
    str总金额医保 = Format(pre_Balance.cur统筹支付 + pre_Balance.cur个人帐户 + pre_Balance.cur公务员补助 + pre_Balance.cur大额统筹 + pre_Balance.cur医院收支 + pre_Balance.cur民政补助 + Val(arrOutput(4)), "#####0.00")
    If str总金额医院 <> str总金额医保 Then
        If Not bln批量预结算 Then
            If MsgBox("医院的费用总金额(" & str总金额医院 & ")与医保中心的费用总额(" & str总金额医保 & ")不等，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            If MsgBox("单病种和精神病因结算政策原因，可能会出现：医院与医保中心的明细总额一致，但结算总额不等" & vbCrLf & _
            "如果点“是”将按单病种及精神病处理，费用差额计入结算方式“医院收支”中；点“否”则按普通病人结算处理！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                bln单病种 = True
            End If
        Else
            strAdvance = "[病人ID:" & lng病人ID & "]HIS总费用与医保总费用不等！" & vbCrLf & _
                "HIS:" & str总金额医保 & Space(5) & "医保:" & str总金额医保
        End If
    End If
    
    住院虚拟结算_重庆 = "医保基金;" & pre_Balance.cur统筹支付 & ";0"
    If pre_Balance.cur个人帐户 <> 0 Then
        住院虚拟结算_重庆 = 住院虚拟结算_重庆 & "|个人帐户;" & pre_Balance.cur个人帐户 & ";0" '不允许修改个人帐户
    End If
    If pre_Balance.cur大额统筹 <> 0 Then
        住院虚拟结算_重庆 = 住院虚拟结算_重庆 & "|大额统筹;" & pre_Balance.cur大额统筹 & ";0"
    End If
    If pre_Balance.cur公务员补助 <> 0 Then
        住院虚拟结算_重庆 = 住院虚拟结算_重庆 & "|公务员补助;" & pre_Balance.cur公务员补助 & ";0"
    End If
    If pre_Balance.cur公务员返还 <> 0 Then
        住院虚拟结算_重庆 = 住院虚拟结算_重庆 & "|公务员返还;" & pre_Balance.cur公务员返还 & ";0"
    End If
    If pre_Balance.cur民政补助 <> 0 Then      '20101028增加民政补助
        住院虚拟结算_重庆 = 住院虚拟结算_重庆 & "|民政补助;" & pre_Balance.cur民政补助 & ";0"
    End If
    
    pre_Balance.cur医院收支 = 0
    If UBound(arrOutput) > 7 Then
        pre_Balance.cur医院收支 = Val(arrOutput(8))
    End If
    If bln单病种 Then
        '接口返回则不需我方计算，否则以HIS总额减去YB总额，差额归入医院收支
        If Not (UBound(arrOutput) > 7) Then
            '需要我方计算（因为接口未返回医院收支，说明还没有这列，因此计算医保总额时不能加医院收支，否则要出错）
            pre_Balance.HIS计算医院收支 = True
            pre_Balance.cur医院收支 = cur发生费用 - (pre_Balance.cur统筹支付 + pre_Balance.cur个人帐户 + pre_Balance.cur公务员补助 + pre_Balance.cur大额统筹 + Val(arrOutput(4)))
        End If
    End If
    If pre_Balance.cur医院收支 <> 0 Then
        住院虚拟结算_重庆 = 住院虚拟结算_重庆 & "|医院收支;" & pre_Balance.cur医院收支 & ";0"
    End If
    
    '保存预结算金额，在结算时再比较一次，避免出现差错
    With g结算数据
        .统筹报销金额 = pre_Balance.cur统筹支付       '1
        .个人帐户支付 = pre_Balance.cur个人帐户       '2
        .累计进入统筹 = pre_Balance.cur公务员补助     '3
        .全自费金额 = Val(arrOutput(4))   '4
        .进入统筹金额 = pre_Balance.cur大额统筹       '5
        .累计统筹报销 = Val(arrOutput(6)) '6
    End With
    
    mlng病人ID = lng病人ID  '表示该病人已经进行了虚拟结算
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_重庆(lng结帐ID As Long, ByVal lng病人ID As Long, Optional ByRef strAdvance As String) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim str操作员 As String, lng结算标志 As Long
    Dim str结算方式 As String
    Dim cur统筹支付 As Double, cur个人帐户 As Double
    Dim cur大额统筹 As Double, cur公务员补助 As Double, cur公务员返还 As Double, cur医院收支 As Double, cur民政补助 As Double, cur民政帐户余额 As Double
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, datCurr As Date, strNO As String
    Dim strFormat As String
    Dim str病种 As String, str并发症 As String, str封锁信息 As String, str备注 As String
    Dim blnOld As Boolean, blnRevise As Boolean '是否需要填写校正字段，是否需要校正结算结果
    Dim bln帐户支付 As Boolean
    Dim int帐户支付方式 As Integer              '0-支付;1-住院询问;2-门诊询问;3-不支付
    
    '单病种增加内容
    Dim str病种编码 As String, bln单病种 As Boolean, int请假 As Integer    '0-未请假中结;1-请假中结中...
    Const int请假开始中结 As Integer = 40
    Const int请假结束中结 As Integer = 50
    
    If mlng病人ID <> lng病人ID Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "该病人没有完成医保的预结算操作，不能进行结算。"
        Exit Function
    End If
    
    On Error GoTo errHandle
    Call DebugTool("进入住院结算")
    '读取参保病人的并发症及封锁信息
    gstrSQL = "Select 退休证号 As 病种编码,病种名称,并发症,封锁信息,Nvl(请假标志,0) AS 请假标志" & _
        " From 保险帐户 " & _
        " Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取参保病人的并发症及封锁信息", lng病人ID, TYPE_重庆市)
    If Not rsTemp.EOF Then
        str病种 = Nvl(rsTemp!病种编码)
        str病种编码 = Nvl(rsTemp!病种编码)
        If str病种 <> "" Then str病种 = "[" & str病种 & "]"
        str病种 = str病种 & Nvl(rsTemp!病种名称)
        str并发症 = Nvl(rsTemp!并发症)
        str封锁信息 = Nvl(rsTemp!封锁信息)
        int请假 = rsTemp!请假标志       '依据保险帐户中的请假标志以判断当前是否在请假中，因此，结算与结算冲销都要设置该标志
    End If
    str备注 = str病种 & "||" & str并发症 & "||" & str封锁信息
    
    '判断是否是单病种结算
    If Trim(str病种编码) <> "" Then
        '打开连接
        If mcnYB.State = 0 Then
            If Not OpenDatabase Then Exit Function
        End If
        gstrSQL = "Select 1 From BZML Where bzfl=5 And upper(BZBM)='" & UCase(str病种编码) & "'"
        Call OpenRecordset_OtherBase(rsTemp, "读取病种分类，供判断是否是单病种", gstrSQL, mcnYB)
        bln单病种 = (rsTemp.RecordCount <> 0)
    End If
    
    '求个人帐户支付金额
    gstrSQL = "Select Nvl(冲预交,0) as 金额 From 病人预交记录 Where 结算方式='个人帐户' And 记录性质=2 And 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "住院结算", lng结帐ID)
    If Not rsTemp.EOF Then cur个人帐户 = rsTemp("金额")
    
    '询问是否使用个人帐户支付
    gstrSQL = "Select Nvl(参数值,0) AS 参数值 From 保险参数 Where 险类=[1] And 参数名='个人帐户'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险参数“离休就诊录入诊断”", TYPE_重庆市)
    If rsTemp.RecordCount <> 0 Then
        int帐户支付方式 = rsTemp!参数值
    End If
    If int帐户支付方式 = 0 Or int帐户支付方式 = 2 Then
        bln帐户支付 = True
    ElseIf int帐户支付方式 = 1 Then
        If MsgBox("请问本次要进行个人帐户支付吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            bln帐户支付 = False
        Else
            bln帐户支付 = True
        End If
    Else
        bln帐户支付 = False
    End If
    
    '调用结算
    With g结算数据
        If .中途结帐 = 1 Then
            '如果病种类型=5（精神病），则提示是否进行请假开始中结
            lng结算标志 = 10
            If bln单病种 Then
                If int请假 = 0 Then
                    If MsgBox("该病人属于单病种，是否进行“请假开始中结”？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        lng结算标志 = int请假开始中结
                        int请假 = 1
                    End If
                Else
                    If MsgBox("该病人属于单病种，本次进行“请假结束中结”，点否则不进行结算？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        lng结算标志 = int请假结束中结
                        int请假 = 0
                    Else
                        Exit Function
                    End If
                End If
            End If
    '            If MsgBox("该病人是否进行转家庭病床结算？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
    '                lng结算标志 = 20 '出院转家庭病床
    '            Else
'                    lng结算标志 = 10 '中途结算
    '            End If
        Else
            lng结算标志 = 0 '正常结算
        End If
        
        StrInput = "05|" & GetIdentify(lng病人ID, .主页ID) & "|" & lng结算标志 & "|" & g结算数据.住院床日 & "|" & UserInfo.姓名 & _
                   "|" & IIf(bln帐户支付, "0", "1") & "|" & Format(mstr费用终止日期, "yyyyMMdd") & "|" & Format(g结算数据.发生费用金额, "#0.00")
    End With
    
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    '填写结算表
    Call DebugTool("填写结算记录")
    datCurr = zlDatabase.Currentdate
    cur个人帐户 = Val(arrOutput(3))
    cur统筹支付 = Val(arrOutput(2))
    cur公务员补助 = Val(arrOutput(4))
    cur大额统筹 = Val(arrOutput(6))
    cur公务员返还 = Val(arrOutput(7))
    If UBound(arrOutput) > 9 Then           '20101028增加民政补助
        cur民政补助 = Val(arrOutput(10))
        cur统筹支付 = cur统筹支付 - cur民政补助
        cur民政帐户余额 = Val(arrOutput(11))
        gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'民政救助门诊余额','''" & cur民政帐户余额 & "''')"      '20101028增加民政补助
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新民政救助门诊余额")
    End If
    If pre_Balance.HIS计算医院收支 = False Then
        If UBound(arrOutput) > 8 Then
            cur医院收支 = Val(arrOutput(9))
        End If
    Else
        '我方再次按规则进行计算
        cur医院收支 = g结算数据.发生费用金额 - (cur个人帐户 + cur统筹支付 + cur公务员补助 + cur大额统筹 + Val(arrOutput(5)))
    End If
    
    '比较正式结算结果是否与虚拟结算结果一致，不一致则需要校正
    If Not (cur个人帐户 = pre_Balance.cur个人帐户 And cur统筹支付 = pre_Balance.cur统筹支付 And _
        cur公务员补助 = pre_Balance.cur公务员补助 And cur大额统筹 = pre_Balance.cur大额统筹 And _
        cur公务员返还 = pre_Balance.cur公务员返还 And cur医院收支 = pre_Balance.cur医院收支 And cur民政补助 = pre_Balance.cur民政补助) Then
        
        blnRevise = True
        
        str结算方式 = "个人帐户|" & cur个人帐户
        If cur统筹支付 <> 0 Then str结算方式 = str结算方式 & "||医保基金|" & cur统筹支付
        If cur大额统筹 <> 0 Then str结算方式 = str结算方式 & "||大额统筹|" & cur大额统筹
        If cur公务员补助 <> 0 Then str结算方式 = str结算方式 & "||公务员补助|" & cur公务员补助
        If cur公务员返还 <> 0 Then str结算方式 = str结算方式 & "||公务员返还|" & cur公务员返还
        If cur医院收支 <> 0 Then str结算方式 = str结算方式 & "||医院收支|" & cur医院收支
        If cur民政补助 <> 0 Then str结算方式 = str结算方式 & "||民政补助|" & cur民政补助
        If str结算方式 <> "" Then
            #If gverControl < 2 Then
                blnOld = True
                gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
            #Else
                strAdvance = str结算方式
                gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
            #End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
        End If
    End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_重庆市, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
    cur帐户增加累计 = mdbl余额
    cur帐户支出累计 = cur个人帐户
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_重庆市 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 & "," & _
        cur进入统筹累计 + cur统筹支付 & "," & _
        cur统筹报销累计 + cur统筹支付 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_重庆市 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,NULL,0," & g结算数据.发生费用金额 & ",0,0," & _
        cur统筹支付 & "," & cur统筹支付 & ",0,0," & cur个人帐户 & ",'" & arrOutput(1) & "'," & g结算数据.主页ID & "," & g结算数据.中途结帐 & ",'" & str备注 & "'" & _
        IIf(blnOld, "", IIf(blnRevise, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    '保险结算计算
    gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",0," & cur统筹支付 & "," & cur统筹支付 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    '出院结算时，将住院门诊号的内容清为空
    If g结算数据.中途结帐 = 0 Then
        Dim str住院门诊号 As String
        If GetMode(lng病人ID, g结算数据.主页ID, str住院门诊号) = False Then
            '将住院门诊号清为空，为下次住院做准备
            gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'住院门诊号','''" & "" & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "将住院门诊号清为空")
        End If
    End If
    
    '更新请假结算标志
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'请假标志','''" & int请假 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新请假中结标志")
    
    #If gverControl < 2 Then
        Call frm结算信息.ShowME(lng结帐ID)
    #End If
    
    住院结算_重庆 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_重庆(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim lng病人ID As Long
    Dim int请假 As Integer
    Dim curDate As Date
        
    On Error GoTo errHandle
    int请假 = 0
    curDate = zlDatabase.Currentdate
    
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", TYPE_重庆市, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    lng病人ID = rsTemp!病人ID
    If Can住院结算冲销(rsTemp("病人ID"), rsTemp("主页ID")) = False Then Exit Function
    
    str流水号 = rsTemp("支付顺序号")
    
    StrInput = "99|" & str流水号 & "|" & ToVarchar(UserInfo.姓名, 20)
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(TYPE_重庆市, rsTemp("病人ID"), Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & rsTemp("病人ID") & "," & TYPE_重庆市 & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - rsTemp("个人帐户支付") & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_重庆市 & "," & rsTemp("病人ID") & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - rsTemp("个人帐户支付") & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & rsTemp("发生费用金额") * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0,0," & _
        rsTemp("个人帐户支付") * -1 & ",'" & str流水号 & "'," & rsTemp("主页ID") & "," & rsTemp("中途结帐") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    '更新请假结算标志
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'请假标志','''" & int请假 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新请假中结标志")

    住院结算冲销_重庆 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 错误信息_重庆(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息

End Function

Public Function 医院编码_重庆() As String
'功能：得到医院的医保编码
    Dim StrInput As String, arrOutput As Variant
    
    On Error GoTo errHandle
    
    StrInput = "11"
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    医院编码_重庆 = arrOutput(1)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function HandleBusiness(ByVal StrInput As String, varOut As Variant, Optional ByVal bln批量预结算 As Boolean = False) As Boolean
'功能：调用医保部件，进行业务处理
    Dim strInfo As String '调用前置服务器的返回值
    Dim lngReturn As Long
    Dim varArray As Variant, lngCount As Long
    
    On Error Resume Next
    varOut = ""
    Screen.MousePointer = vbHourglass
    strInfo = Space(1024)
    lngReturn = dy_Business_Handle(StrInput, strInfo)
'    lngReturn = gobj重庆市东软.dy_Business_Handle(StrInput, strInfo)
    If Err <> 0 Or lngReturn = -1 Then
        varArray = Split(strInfo, "|")
        
        If UBound(varArray) > 0 Then
            strInfo = "医保接口调用失败。" & vbCrLf & varArray(1)
        Else
            strInfo = "医保接口调用失败。" & vbCrLf & strInfo
        End If
        Screen.MousePointer = vbDefault
        If Not bln批量预结算 Then MsgBox strInfo, vbExclamation, gstrSysName
        Exit Function
    End If
    strInfo = TruncZero(strInfo)
    
    varArray = Split(strInfo, "|")
    If varArray(0) = "-1" Then
        '业务调用失败
        If UBound(varArray) > 0 Then
            strInfo = "医保接口出现警告。" & vbCrLf & varArray(1)
        Else
            strInfo = "医保业务处理失败。"
        End If
        
        Screen.MousePointer = vbDefault
        If Not bln批量预结算 Then MsgBox strInfo, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '交易成功
    varOut = Split(strInfo, "|")
    
    HandleBusiness = True
    Screen.MousePointer = vbDefault
End Function

Private Function Get保险参数_重庆(ByVal str参数名 As String) As String
'功能：获得保险参数
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.参数名,A.参数值 from 保险参数 A " & _
              " where A.参数名=[1] and A.险类=[2] and A.中心 is null "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", str参数名, TYPE_重庆市)
    
    If rsTemp.EOF = False Then
        Get保险参数_重庆 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
    End If
End Function

Public Function 价格判断_重庆(ByVal dbl医院 As Double, ByVal dbl医保 As Double, ByVal str限价方式 As String, _
                              ByVal bln特价 As Boolean, ByVal dbl特价 As Double) As Boolean
'功能：判断医院的价格是否超过医保规定的单价
    Dim str医院类别 As String
    
    On Error GoTo errHandle
    
    If InStr(str限价方式, "二级") > 0 Then
        str医院类别 = Get保险参数_重庆("医院等级")
        '给出的标准价格为二级医院的最高限价，三级医院的最高限价在此基础上可以上浮10%，一级医院的最高限价在此基础上下调5%
        
        Select Case str医院类别
            Case "三级"
                dbl医保 = dbl医保 * 1.1
            Case "一级"
                dbl医保 = dbl医保 * 0.95
        End Select
    End If
    
    If bln特价 = True And dbl特价 > dbl医保 Then
        '允许使用特价
        dbl医保 = dbl特价
    End If
    
    If dbl医院 > dbl医保 Then
        If MsgBox("医院单价" & Format(dbl医院, "0.000") & " 高于医保中心核准的价格" & Format(dbl医保, "0.000") & "，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    价格判断_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 记帐传输_重庆(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0) As Boolean
'功能:上传新产生的记帐明细到医保中心
'参数:  str单据号   NO
'       int性质     记录性质
'       str消息    如果传输过程中有提醒，传回前台程序完成（避免长时间的死锁）
'       lng病人ID  默认为0，表示传输整张单据，否则为单据中指定病人的。（主要是因为医嘱在保存记帐单时，是分病人在提交数据而不是一起提交）
'返回:
    Dim rsTemp As New ADODB.Recordset
    Dim rsTest As New ADODB.Recordset
    Dim cn上传 As New ADODB.Connection
    Dim StrInput As String, arrOutput, arrTemp  As Variant, cur统筹金额 As Currency
    Dim str医生 As String, str经办人 As String
    Dim col病人 As New Collection, lngPre病人ID As Long, var病人 As Variant, bln成功 As Boolean
    Dim str冲销明细流水号 As String, str封锁信息 As String, str负数冲销信息 As String
    '请注意：重庆医保是在记帐单保存后再调用传输过程的。
    
    On Error GoTo errHandle
    
    Set cn上传 = GetNewConnection
    cn上传.Open
    
    '读出该张单据的费用明细
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.主页ID,A.发生时间 as 登记时间,Round(A.实收金额,4) 实收金额 " & _
              "         ,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         ,C.项目编码,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,E.规格,G.名称 剂型 " & _
              "  From 住院费用记录 A,收费细目 B,保险支付项目 C,病案主页 D,药品目录 E ,药品信息 F,药品剂型 G " & _
              "  where A.NO=[1] and A.记录性质=[2] and A.记录状态=1 And Nvl(A.是否上传,0)=0 And Nvl(A.实收金额,0)<>0 " & _
              "        and A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=" & TYPE_重庆市 & IIf(lng病人ID = 0, "", " and A.病人ID=[3]") & _
              "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类=D.险类 " & _
              "        AND B.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) " & _
              "  Order by A.病人ID,A.发生时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "记帐传输", str单据号, int性质, lng病人ID)
    
    '以前的检查不需要，本来就是保存后上传，检查也没有意义
'    Do While Not rsTemp.EOF
'        If Val(rsTemp!数量) < 0 Or Val(rsTemp!价格) < 0 Then
'            '任意取一笔正常记录的流水号，做为冲销流水号
'            str冲销明细流水号 = GetSequence(rsTemp!病人ID, rsTemp!主页ID, rsTemp!收费细目ID)
'            If Trim(str冲销明细流水号) = "" Then
'                MsgBox "没有找到可以冲销的记录！[" & rsTemp!编码 & "]" & rsTemp!名称, vbInformation, gstrSysName
'                Exit Function
'            End If
'        Else
'            str冲销明细流水号 = ""
'        End If
'        rsTemp.MoveNext
'    Loop
'    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
'
'    '保存医保项目的信息，备查
'    Do While Not rsTemp.EOF
'        Call TrackRecordInsure(rsTemp!ID, rsTemp!收费细目ID)
'        rsTemp.MoveNext
'    Loop
    
    '进行费用明细的传输
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do Until rsTemp.EOF
        If Val(rsTemp!数量) < 0 Or Val(rsTemp!价格) < 0 Then
            '任意取一笔正常记录的流水号，做为冲销流水号
            str冲销明细流水号 = "冲销记录"
        Else
            str冲销明细流水号 = ""
        End If
        
        If str冲销明细流水号 = "" Then
            Call DebugTool("准备上传正常记帐明细")
            str医生 = ToVarchar(IIf(IsNull(rsTemp("医生")), UserInfo.姓名, rsTemp("医生")), 20)
            str经办人 = ToVarchar(IIf(IsNull(rsTemp("操作员姓名")), UserInfo.姓名, rsTemp("操作员姓名")), 20)
            
            StrInput = "04|" & GetIdentify(rsTemp("病人ID"), rsTemp("主页ID"))
            StrInput = StrInput & "|" & rsTemp("NO") & "_" & int性质
            StrInput = StrInput & "|" & Format(rsTemp("登记时间"), "yyyy-MM-dd HH:mm:ss")
            StrInput = StrInput & "|" & ToVarchar(rsTemp("项目编码"), 10)     '中心编码
            StrInput = StrInput & "|" & ToVarchar(rsTemp("编码"), 20)         '医院内码
            StrInput = StrInput & "|" & ToVarchar(rsTemp("名称"), 50)         '项目名称
            StrInput = StrInput & "|" & Format(rsTemp("价格"), "0.0000")      '单价
            StrInput = StrInput & "|" & Format(rsTemp("数量"), "0.00")        '数量
            StrInput = StrInput & "|" & IIf(rsTemp("是否急诊") = 1, 1, 0)     '急诊标志
            StrInput = StrInput & "|" & str医生                               '医生
            StrInput = StrInput & "|" & str经办人                             '经办人
            StrInput = StrInput & "|" & ToVarchar(rsTemp("计算单位"), 20)     '单位
            StrInput = StrInput & "|" & ToVarchar(rsTemp("规格"), 14)         '规格
            StrInput = StrInput & "|" & ToVarchar(rsTemp("剂型"), 20)         '剂理
            StrInput = StrInput & "|" & str冲销明细流水号                     '冲销明细流水号
            StrInput = StrInput & "|" & Format(rsTemp("实收金额"), "#####0.0000")         '金额
            
            If HandleBusiness(StrInput, arrOutput) = False Then
                If bln成功 = True Then
                    MsgBox "数据上传中途发生错误，并且已经部分已经上传，请在预结算处完成剩余数据的上传。", vbInformation, gstrSysName
                Else
                    MsgBox "数据上传发生错误，没有成功上传的记录，请在预结算处完成剩余数据的上传。", vbInformation, gstrSysName
                End If
                记帐传输_重庆 = True
                Exit Function
            End If
            Call AddMessage(str消息, arrOutput, rsTemp("名称"), rsTemp("价格"))  '可以产生的提醒信息
            
            '在费用记录上打上标记，说明已经上传，并保存返回的金额
            If arrOutput(3) = 2 Then
                '未通过审批
                cur统筹金额 = 0
            Else
                '特准单价 * 数量
                cur统筹金额 = Val(arrOutput(2)) * rsTemp("数量")
            End If
            
            '正常记录上传，仅记录处方流水号
            gstrSQL = "ZL_病人记帐记录_上传(" & rsTemp("ID") & "," & cur统筹金额 & ",'" & arrOutput(1) & "')"
            cn上传.Execute gstrSQL, , adCmdStoredProc
            bln成功 = True
        Else
            str负数冲销信息 = UploadNegative(rsTemp!病人ID, rsTemp!主页ID, rsTemp!ID, rsTemp!收费细目ID)
            If str负数冲销信息 <> "" Then str消息 = str消息 & str负数冲销信息 & vbCrLf
            bln成功 = (str负数冲销信息 = "")
        End If
        
        If lngPre病人ID <> rsTemp("病人ID") Then '判断时没有考虑主页ID，是因为同一病人不可能同时有两次住院的明细
            'Modified by ZYB 2004-05-10
            '提取病人的基本信息，如果存在封锁原因则提示
            Call DebugTool("获取病人封锁信息并更新")
            gstrSQL = "Select 医保号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
            Set rsTest = zlDatabase.OpenSQLRecord(gstrSQL, "取该病人的医保号", TYPE_重庆市, CLng(rsTemp!病人ID))
            
            StrInput = "01|" & rsTest!医保号
            If HandleBusiness(StrInput, arrTemp) Then
                str封锁信息 = ""
                If Val(arrTemp(11)) <> 0 Then
                    str封锁信息 = arrTemp(12)
                    MsgBox str封锁信息, vbInformation, gstrSysName
                End If
                '更新封锁信息
                gstrSQL = "ZL_保险帐户_更新信息(" & rsTemp!病人ID & "," & TYPE_重庆市 & ",'封锁信息','''" & str封锁信息 & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新封锁信息")
            End If
            
            '将已经上传的病人信息记录下来（因为记帐表是多病人的）
            col病人.Add rsTemp("病人ID") & "_" & rsTemp("主页ID")
            lngPre病人ID = rsTemp("病人ID")
        End If
        
        rsTemp.MoveNext
    Loop
    
    If str消息 <> "" Then
        str消息 = "病人费用明细传输过程中得到医保中心如下反馈信息，但目前数据已经保存。" & vbCrLf & "如果有何不妥，你可以选择作废该单据。" & vbCrLf & vbCrLf & str消息
    End If
        
    记帐传输_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 记帐作废_重庆(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String) As Boolean
'功能:作废已经上传到医保中心的记帐明细
'参数:  str单据号   NO
'       int性质     记录性质
'       str消息    如果传输过程中有提醒，传回前台程序完成（避免长时间的死锁）
'返回:
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, arrOutput As Variant
    Dim lngPre病人ID As Long
    Dim str医生 As String, str经办人 As String, str冲销明细流水号 As String
    Dim bln成功 As Boolean
    Dim cn上传 As New ADODB.Connection
    
    On Error GoTo errHandle
    
    Set cn上传 = GetNewConnection
    cn上传.Open
    
    '读出该张单据的费用明细中有未上传的记录（取原始单据）
    gstrSQL = "Select nvl(count(A.ID),0) as 总数,nvl(sum(A.是否上传),0) 上传数 " & _
              "  From 住院费用记录 A,病案主页 B,保险支付项目 C" & _
              "  where A.NO=[1] and A.记录性质=[3] and A.记录状态<>2 And Nvl(A.记录状态,0)<>0 and nvl(A.实收金额,0)<>0  " & _
              "        and A.病人ID=B.病人ID and A.主页ID=B.主页ID And B.险类=[1] and A.收费细目ID=C.收费细目ID and B.险类=C.险类"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "记帐作废", TYPE_重庆市, str单据号, int性质)
    
    If rsTemp.EOF = True Then
        MsgBox "该单据里没有可上传的作废费用明细。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rsTemp("上传数") = 0 Then
        '明细根本就没有上传，所以也就不需要处理作废
        记帐作废_重庆 = True
        Exit Function
    End If
    
    If rsTemp("上传数") < rsTemp("总数") Then
        MsgBox "该单据里还有未上传的费用明细，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '读出该单据内病人情况（因为记帐表是多病人的）
    gstrSQL = " Select A.ID,A.收费细目ID,A.NO,A.记录性质,A.记录状态,A.序号,A.病人ID,A.主页ID,A.发生时间 as 登记时间,Round(A.实收金额,4) 实收金额" & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         ,C.项目编码,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,E.规格,G.名称 剂型 " & _
              "  From 住院费用记录 A,收费细目 B,保险支付项目 C,病案主页 D,药品目录 E ,药品信息 F,药品剂型 G " & _
              "  where A.NO=[1] and A.记录性质=[2] and A.记录状态=2 and nvl(A.是否上传,0)=0 And Nvl(A.实收金额,0)<>0" & _
              "        and A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=[3]" & _
              "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类=D.险类 " & _
              "        AND B.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) " & _
              "  Order by A.发生时间,A.记录性质,Decode(A.记录状态,2,2,1)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "记帐作废", str单据号, int性质, TYPE_重庆市)
    
    '先将医保项目信息保存起来，备查
    Do While Not rsTemp.EOF
        Call TrackRecordInsure(rsTemp!ID, rsTemp!收费细目ID)
        rsTemp.MoveNext
    Loop
    
    '进行费用明细的传输
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do Until rsTemp.EOF
        '记帐表和记帐单都允许部分冲销后，需要一笔一笔冲销
        str冲销明细流水号 = GetDetailSequence(rsTemp!NO, rsTemp!序号, rsTemp!记录性质, rsTemp!记录状态)
        str经办人 = ToVarchar(IIf(IsNull(rsTemp("操作员姓名")), UserInfo.姓名, rsTemp("操作员姓名")), 20)
        StrInput = "99|" & str冲销明细流水号 & "|" & str经办人
        If HandleBusiness(StrInput, arrOutput) = False Then
            '费用上传失败
            If bln成功 = True Then
                MsgBox "数据上传中途发生错误，并且已经部分已经上传，请在预结算处完成剩余数据的上传。", vbInformation, gstrSysName
            Else
                MsgBox "数据上传发生错误，没有成功上传的记录，请在预结算处完成剩余数据的上传。", vbInformation, gstrSysName
            End If
            记帐作废_重庆 = True
            Exit Function
        Else
            '在产生的作废费用记录上打上标记，说明已经上传
            gstrSQL = "ZL_病人记帐记录_上传(" & rsTemp("ID") & ")"
            '采用另一个连接串执行，已成功上传的打上上传标志
            cn上传.Execute gstrSQL, , adCmdStoredProc
        End If
        
        rsTemp.MoveNext
        bln成功 = True
    Loop
    
    记帐作废_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub AddMessage(strMessage As String, arrOutput As Variant, ByVal str项目 As String, ByVal dbl单价 As Currency, Optional ByVal bln住院 As Boolean = True)
'功能：在病人费用明细传输过程中可能产生一些需要提醒操作人员的信息
    Dim strTemp As String
    
    If dbl单价 > Val(arrOutput(2)) And Val(arrOutput(2)) > 0 Then
        strTemp = "●    " & str项目 & "的医院价格 " & Format(dbl单价, "0.0000") & " 高于中心返回价格 " & Format(Val(arrOutput(2)), "0.0000") & vbCrLf
    End If
    If arrOutput(3) = 2 And bln住院 Then
        strTemp = "●    " & str项目 & "需要审批，但没有审批记录，只能作为自费项目" & vbCrLf
    End If
    
    If InStr(strMessage, strTemp) = 0 Then
        strMessage = strMessage & strTemp
    End If
    
End Sub

'摘要的格式：（正常记录）原始流水号|已冲销数量；（冲销记录）原始流水号|被冲销流水号
Private Function GetDetailSequence(ByVal strNO As String, ByVal int序号 As Integer, _
        ByVal int性质 As Integer, ByVal int状态 As Integer) As String
    Dim rsTemp As New ADODB.Recordset
    '冲销时使用：根据NO、记录性质、记录状态、序号取原始记录的流水号
    GetDetailSequence = ""
    If int状态 <> 2 Then Exit Function
    
    gstrSQL = " Select 摘要 From 住院费用记录" & _
              " Where NO=[1] And 序号=[2]" & _
              " And 记录性质=[3] And 记录状态=3"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取原始处方明细的流水号", strNO, int序号, int性质)
    If Not rsTemp.EOF Then
        GetDetailSequence = Split(Nvl(rsTemp!摘要, "|"), "|")(0)
    Else
        Call DebugTool("未找到原始处方明细[NO:" & strNO & "|序号:" & int序号 & "|记录性质:" & int性质 & "|记录状态:" & int状态)
    End If
End Function

Private Function UploadNegative(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngID As Long, ByVal lng收费细目ID As Long) As String
    '负数记帐时使用：规则为a)先取数量足够的；b)再取明细累计足够的
    '现统计，现冲销，现保存数据，现更新已冲销数量
    Dim arrOutput
    Dim cur统筹金额 As Currency, cur统筹金额累计 As Currency    '一个用于完整冲销，一个用于分笔冲销记录累计值
    Dim StrInput As String
    Dim str医生 As String, str经办人 As String
    
    Dim str摘要 As Double     '数量,流水号|数量,流水号，用于保存明细记录时使用
    Dim str被冲销流水号 As String, str流水号 As String
    Dim dbl冲销数量 As Double, dbl待冲销总数 As Double, dbl已冲销数量 As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rsSource As New ADODB.Recordset
    Dim rsFilter As New ADODB.Recordset
    On Error GoTo errHand
    
    '读取本次待冲销总数
    gstrSQL = "Select A.ID,A.NO,A.记录性质,A.记录状态,A.病人ID,A.主页ID,A.发生时间 as 登记时间,Round(A.实收金额,4) 实收金额 " & _
              "         ,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         ,C.项目编码,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,E.规格,G.名称 剂型 " & _
              "  From 住院费用记录 A,收费细目 B,保险支付项目 C,病案主页 D,药品目录 E ,药品信息 F,药品剂型 G " & _
              "  where A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=[1]" & _
              "        And A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类=D.险类 And Nvl(A.是否上传,0)=0" & _
              "        AND B.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) And A.ID=[2]"
    Set rsSource = zlDatabase.OpenSQLRecord(gstrSQL, "读取本次待冲销总数", TYPE_重庆市, lngID)
    If rsSource.RecordCount <> 0 Then
        dbl待冲销总数 = Abs(rsSource!数量)
    End If
    Call DebugTool("读取本次待冲销总数")
    
    '读取本记录已冲销的数量
    gstrSQL = " Select SUM(数量) AS 已冲销数量 From 负数冲销明细 Where 费用ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取本记录已冲销的数量", lngID)
    dbl待冲销总数 = dbl待冲销总数 - Nvl(rsTemp!已冲销数量, 0)
    '对于已冲销完成但未更新上传标志的处理
    If dbl待冲销总数 = 0 Then Exit Function
    
    '读取存在剩余数量可冲销的原始记录（取摘要中记录的已冲销数量）
    gstrSQL = " Select A.ID,A.付数*A.数次 AS 数量,A.摘要," & _
              "     To_Number(Nvl(Substr(A.摘要,Decode(Instr(A.摘要,'|',1,1),0,Length(A.摘要),Instr(A.摘要,'|',1,1))+1),0)) as 已冲销数量" & _
              " From 住院费用记录 A" & _
              " Where A.记录状态=1 And Nvl(A.实收金额,0)<>0 And Nvl(A.附加标志,0)<>9 " & _
              " And Nvl(A.是否上传,0)=1 And A.收费细目ID=" & lng收费细目ID & _
              " And A.病人ID=[1] And A.主页ID=[2]"
    gstrSQL = " Select ID,数量-已冲销数量 AS 剩余数量,数量,已冲销数量,摘要 From (" & gstrSQL & ") Where 数量-已冲销数量>0 Order by 剩余数量"
    Set rsFilter = zlDatabase.OpenSQLRecord(gstrSQL, "读取存在剩余数量可冲销的原始记录", lng病人ID, lng主页ID)
    Call DebugTool("读取存在剩余数量可冲销的原始记录")
    
    '按一次性匹配原则，取大于等于本次冲销数量的记录，符合就马上退出
    With rsFilter
        .Filter = "剩余数量>=" & dbl待冲销总数
        If .RecordCount <> 0 Then
            dbl冲销数量 = dbl待冲销总数
            str被冲销流水号 = Split(Nvl(!摘要, "|"), "|")(0)
            
            '上传负数记录
            str医生 = ToVarchar(IIf(IsNull(rsSource("医生")), UserInfo.姓名, rsSource("医生")), 20)
            str经办人 = ToVarchar(IIf(IsNull(rsSource("操作员姓名")), UserInfo.姓名, rsSource("操作员姓名")), 20)
            
            StrInput = "04|" & GetIdentify(rsSource("病人ID"), rsSource("主页ID"))
            StrInput = StrInput & "|" & rsSource("NO") & "_" & rsSource!记录性质
            StrInput = StrInput & "|" & Format(rsSource("登记时间"), "yyyy-MM-dd HH:mm:ss")
            StrInput = StrInput & "|" & ToVarchar(rsSource("项目编码"), 10)     '中心编码
            StrInput = StrInput & "|" & ToVarchar(rsSource("编码"), 20)         '医院内码
            StrInput = StrInput & "|" & ToVarchar(rsSource("名称"), 50)         '项目名称
            StrInput = StrInput & "|" & Format(rsSource("价格"), "0.0000")      '单价
            StrInput = StrInput & "|" & Format(-1 * dbl冲销数量, "0.00")       '数量
            StrInput = StrInput & "|" & IIf(rsSource("是否急诊") = 1, 1, 0)     '急诊标志
            StrInput = StrInput & "|" & str医生                               '医生
            StrInput = StrInput & "|" & str经办人                             '经办人
            StrInput = StrInput & "|" & ToVarchar(rsSource("计算单位"), 20)     '单位
            StrInput = StrInput & "|" & ToVarchar(rsSource("规格"), 14)         '规格
            StrInput = StrInput & "|" & ToVarchar(rsSource("剂型"), 20)         '剂理
            StrInput = StrInput & "|" & str被冲销流水号                     '冲销明细流水号
            StrInput = StrInput & "|" & Format(-1 * rsSource("价格") * dbl冲销数量, "#0.0000")     '金额
            
            Call DebugTool("准备调用上传，被冲销流水号:" & str被冲销流水号 & ";冲销数量:" & dbl冲销数量 & ";待冲销数量:" & dbl待冲销总数)
            If HandleBusiness(StrInput, arrOutput) = False Then
                UploadNegative = "费用ID=" & lngID & "的负数冲销记录上传失败！（被冲销记录流水号=" & str被冲销流水号 & "；冲销数量=" & dbl冲销数量 & "）"
                Exit Function
            Else
                '成功就保存此明细，并更新已冲销数量
                '在费用记录上打上标记，说明已经上传，并保存返回的金额
                If arrOutput(3) = 2 Then
                    '未通过审批
                    cur统筹金额 = 0
                Else
                    '特准单价 * 数量
                    cur统筹金额 = Val(arrOutput(2)) * rsSource("数量")
                End If
                '更新原始处方的已冲销数量
                dbl已冲销数量 = !已冲销数量 + dbl冲销数量
                gstrSQL = "ZL_病人记帐记录_上传(" & !ID & ",NULL,'" & str被冲销流水号 & "|" & dbl已冲销数量 & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                '打上传标志
                gstrSQL = "ZL_病人记帐记录_上传(" & lngID & "," & cur统筹金额 & ",'" & arrOutput(1) & "|" & str被冲销流水号 & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                '产生冲销明细
                gstrSQL = "zl_负数冲销明细_Insert(" & lng病人ID & "," & lng主页ID & "," & lngID & "," & dbl冲销数量 & ",'" & arrOutput(1) & "','" & str被冲销流水号 & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                Call DebugTool("更新上传标志，被冲销明细剩余数量，并产生冲销明细记录")
            End If
            dbl待冲销总数 = 0
        End If
        .Filter = 0
        
        If dbl待冲销总数 <> 0 Then
            '说明无足够的数量可冲销，查剩余累计是否存在足够冲销的
            Do While Not .EOF
                If dbl待冲销总数 > !剩余数量 Then
                    dbl冲销数量 = !剩余数量
                Else
                    dbl冲销数量 = dbl待冲销总数
                End If
                str被冲销流水号 = Split(Nvl(!摘要, "|"), "|")(0)
                
                '上传负数记录
                str医生 = ToVarchar(IIf(IsNull(rsSource("医生")), UserInfo.姓名, rsSource("医生")), 20)
                str经办人 = ToVarchar(IIf(IsNull(rsSource("操作员姓名")), UserInfo.姓名, rsSource("操作员姓名")), 20)
                
                StrInput = "04|" & GetIdentify(rsSource("病人ID"), rsSource("主页ID"))
                StrInput = StrInput & "|" & rsSource("NO") & "_" & rsSource!记录性质
                StrInput = StrInput & "|" & Format(rsSource("登记时间"), "yyyy-MM-dd HH:mm:ss")
                StrInput = StrInput & "|" & ToVarchar(rsSource("项目编码"), 10)     '中心编码
                StrInput = StrInput & "|" & ToVarchar(rsSource("编码"), 20)         '医院内码
                StrInput = StrInput & "|" & ToVarchar(rsSource("名称"), 50)         '项目名称
                StrInput = StrInput & "|" & Format(rsSource("价格"), "0.0000")      '单价
                StrInput = StrInput & "|" & Format(-1 * dbl冲销数量, "0.00")       '数量
                StrInput = StrInput & "|" & IIf(rsSource("是否急诊") = 1, 1, 0)     '急诊标志
                StrInput = StrInput & "|" & str医生                               '医生
                StrInput = StrInput & "|" & str经办人                             '经办人
                StrInput = StrInput & "|" & ToVarchar(rsSource("计算单位"), 20)     '单位
                StrInput = StrInput & "|" & ToVarchar(rsSource("规格"), 14)         '规格
                StrInput = StrInput & "|" & ToVarchar(rsSource("剂型"), 20)         '剂理
                StrInput = StrInput & "|" & str被冲销流水号                     '冲销明细流水号
                StrInput = StrInput & "|" & Format(-1 * rsSource("价格") * dbl冲销数量, "#0.0000")     '金额
                
                If HandleBusiness(StrInput, arrOutput) = False Then
                    UploadNegative = "费用ID=" & lngID & "的负数冲销记录上传失败！（被冲销记录流水号=" & str被冲销流水号 & "；冲销数量=" & dbl冲销数量 & "）"
                    Exit Function
                Else
                    '成功就保存此明细，并更新已冲销数量
                    '在费用记录上打上标记，说明已经上传，并保存返回的金额
                    If arrOutput(3) = 2 Then
                        '未通过审批
                        cur统筹金额 = 0
                    Else
                        '特准单价 * 数量
                        cur统筹金额 = -1 * Val(arrOutput(2)) * dbl冲销数量
                    End If
                    cur统筹金额累计 = cur统筹金额累计 + cur统筹金额
                    dbl已冲销数量 = !已冲销数量 + dbl冲销数量
                    dbl待冲销总数 = dbl待冲销总数 - dbl冲销数量
                    
                    '更新原始处方的已冲销数量
                    gstrSQL = "ZL_病人记帐记录_上传(" & !ID & ",NULL,'" & str被冲销流水号 & "|" & dbl已冲销数量 & "')"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    '产生冲销明细
                    gstrSQL = "zl_负数冲销明细_Insert(" & lng病人ID & "," & lng主页ID & "," & lngID & "," & dbl冲销数量 & ",'" & arrOutput(1) & "','" & str被冲销流水号 & "')"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    Call DebugTool("更新被冲销明细剩余数量，并产生冲销明细记录")
                End If
                
                If dbl待冲销总数 = 0 Then
                    '打上传标志
                    gstrSQL = "ZL_病人记帐记录_上传(" & lngID & "," & cur统筹金额累计 & ",'" & arrOutput(1) & "|" & str被冲销流水号 & "')"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    '已冲销完毕，正常结束
                    Call DebugTool("更新上传标志")
                    Exit Function
                End If
                
                .MoveNext
            Loop
        End If
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function TrackRecordInsure(ByVal lng费用ID As Long, ByVal lng收费细目ID As Long) As Boolean
    Dim str流水号 As String, str费用类型 As String, str收费类别 As String, str备注 As String
    Dim dbl标准单价 As Double, dbl自付比例 As Double
    Dim rsTemp As New ADODB.Recordset

    '20061113,zyb:取消TrackRecordInsure()的调用，避免打开医保前置机连接数过多（东软限制了连接数，调dy_init也会打开一个连接）
    Exit Function
    
    '记录医保项目此时的基本信息（医保项目编码，费用类型，标准单价,自付比例）
    Call DebugTool("进入TrackRecordInsure")
    gstrSQL = "Select A.类别,B.项目编码 " & _
        " From 收费细目 A,保险支付项目 B" & _
        " Where A.ID=B.收费细目ID And B.险类=[1] And A.ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取该项目所属的类别", TYPE_重庆市, lng收费细目ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    str流水号 = Nvl(rsTemp!项目编码)
    str收费类别 = rsTemp!类别
    Call DebugTool("当前医保项目编码：" & str流水号)
    If str流水号 = "" Then Exit Function
    
    '打开连接
    If mcnYB.State = 0 Then
        If Not OpenDatabase Then Exit Function
    End If
    
    If InStr(1, "5,6,7", str收费类别) <> 0 Then
        '药品
        gstrSQL = "select YPLSH  医保编码,YPBM 药品编码,TYM 通用名称,SPM 商品名,SPMZJM 商品名简码,YCMC 药厂名称,decode(FYDJ,1,'甲类',2,'乙类','自费') 费用等级 " & _
                  "      ,PFJ 批发价,BZDJ 标准单价,ZFBL 自付比例,JX 剂型,BZSL 包装数量,BZDW 包装单位,HL 含量,HLDW 含量单位,RL 容量,RLDW 容量单位 " & _
                  "      ,DECODE(CFYBZ,1,'是') 处方药标志,decode(GMP,1,'是') GMP标志,decode(YPXJFS,1,'限价') 限价,TQFYDJ 特群项目等级,TQZFBL 特群自付比例,TQBZDJ 特群标准单价 " & _
                  "  FROM YPML WHERE YPLSH='" & str流水号 & "'"
    Else
        '诊疗
        gstrSQL = "Select XMLSH 医保编码,XMBM 诊疗编码,XMMC 项目名称,ZJM 简码,decode(FYDJ,1,'甲类',2,'乙类','自费') 费用等级,DW 单位 " & _
                 "       ,TPJ 特批价,BZJ 标准单价,ZZBL 在职自付比例,TXBL 退休自付比例,decode(XJFS,1,'统一限价',2,'按医院等级定价',3,'按二级医院标准浮动比例') 限价 " & _
                 "       ,TQFYDJ 特群项目等级,TQZFBL 特群自付比例,TQBZDJ 特群标准单价,decode(TPXMBZ,1,'是') 特批项目标志,BZ 备注 " & _
                 "   FROM ZLXM WHERE XMLSH='" & str流水号 & "'"
    End If
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, mcnYB
        If .RecordCount = 0 Then
            Call DebugTool("未找到该医保项目")
            Exit Function
        End If
    End With
    
    str费用类型 = Nvl(rsTemp!费用等级)
    dbl标准单价 = Nvl(rsTemp!标准单价, 0)
    If InStr(1, "5,6,7", str收费类别) <> 0 Then
        dbl自付比例 = Nvl(rsTemp!自付比例, 0)
        str备注 = "||||" & Nvl(rsTemp!特群自付比例, 0)
    Else
        dbl自付比例 = Nvl(rsTemp!在职自付比例, 0)
        str备注 = Nvl(rsTemp!特批价, 0) & "||" & Nvl(rsTemp!退休自付比例, 0) & "||" & Nvl(rsTemp!特群自付比例, 0)
    End If
    
    '插入记录（过程中判断，如果存在记录则仅不更新，否则插入）
    '费用ID,医保项目编码,费用类型,标准单价,自付比例,备注
    gstrSQL = "zl_医保项目信息_INSERT(" & lng费用ID & ",'" & str流水号 & "','" & str费用类型 & "'," & _
        dbl标准单价 & "," & dbl自付比例 & ",'" & str备注 & "')"
    Call DebugTool("插入医保项目信息：" & gstrSQL)
    Call zlDatabase.ExecuteProcedure(gstrSQL, "插入医保项目信息记录")
    TrackRecordInsure = True
End Function

Private Function OpenDatabase() As Boolean
    Dim strServer As String, strUser As String, strPass As String, strTemp As String
    Dim rsTemp As New ADODB.Recordset
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取保险参数", TYPE_重庆市)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保服务器"
                strServer = strTemp
            Case "医保用户名"
                strUser = strTemp
            Case "医保用户密码"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
        Exit Function
    End If
    OpenDatabase = True
End Function

Public Function 挂号结算_重庆(lng结帐ID As Long) As Boolean
    Dim intTotal As Integer, intStart As Integer
    Dim str结算方式 As String, arr结算方式
    Dim cur个人帐户 As Currency, cur医保基金 As Currency, cur公务员补助 As Currency, cur大额统筹 As Currency, cur公务员返还 As Currency, cur民政补助 As Currency
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    gstrSQL = "Select 病人ID,收费细目ID,数次*NVL(付数,1) AS 数量,标准单价 As 单价,实收金额,开单人," & IIf(g结算数据.超限自付金额 = 14, "1", "0") & " As 是否急诊" & _
        " From 门诊费用记录 " & _
        " Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", lng结帐ID)
    
    If Not 门诊虚拟结算_重庆(rsTemp, str结算方式) Then Exit Function
    If Not 门诊结算_重庆(lng结帐ID, 0, "", "", False) Then Exit Function
    
    '分解各种结算方式
    arr结算方式 = Split(str结算方式, "|")
    intTotal = UBound(arr结算方式)
    For intStart = 0 To intTotal
        Select Case Split(arr结算方式(intStart), ";")(0)
        Case "个人帐户"
            cur个人帐户 = Val(Split(arr结算方式(intStart), ";")(1))
        Case "医保基金"
            cur医保基金 = Val(Split(arr结算方式(intStart), ";")(1))
        Case "公务员补助"
            cur公务员补助 = Val(Split(arr结算方式(intStart), ";")(1))
        Case "大额统筹"
            cur大额统筹 = Val(Split(arr结算方式(intStart), ";")(1))
        Case "公务员返还"
            cur公务员返还 = Val(Split(arr结算方式(intStart), ";")(1))
                Case "民政补助"
            cur民政补助 = Val(Split(arr结算方式(intStart), ";")(1))
        End Select
    Next
    
   '需要修正结算结果
    str结算方式 = ""
    If cur个人帐户 <> 0 Then str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户
    If cur医保基金 <> 0 Then str结算方式 = str结算方式 & "||医保基金|" & cur医保基金
    If cur大额统筹 <> 0 Then str结算方式 = str结算方式 & "||大额统筹|" & cur大额统筹
    If cur公务员补助 <> 0 Then str结算方式 = str结算方式 & "||公务员补助|" & cur公务员补助
    If cur公务员返还 <> 0 Then str结算方式 = str结算方式 & "||公务员返还|" & cur公务员返还
        If cur民政补助 <> 0 Then str结算方式 = str结算方式 & "||民政补助|" & cur民政补助
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
    Else
        str结算方式 = "个人帐户|0"
    End If
    gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
    
    挂号结算_重庆 = True
    Call frm结算信息.ShowME(lng结帐ID)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetIdentify(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal bln更新 As Boolean = False) As String
    Dim str住院门诊号 As String
    'MODIFIED BY ZYB 20040626 '三院修改，由于三晶老系统是按病人ID上传的，因此，流水号改为：
    '如果保险帐户中存在住院门诊号，且值不为空，则以它做为病人标识上传；否则按现有模式上传
    If GetMode(lng病人ID, lng主页ID, str住院门诊号) Then
        GetIdentify = lng病人ID & "_" & lng主页ID & "_" & Get就诊序号(lng病人ID, bln更新)
    Else
        GetIdentify = str住院门诊号
    End If
End Function

Private Function Get就诊序号(ByVal lng病人ID As Long, ByVal bln更新 As Boolean) As Integer
    Dim int序号 As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    gstrSQL = " Select 就诊序号 From 保险帐户 " & _
              " Where 险类=[1] ANd 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取住院门诊号", TYPE_重庆市, lng病人ID)
    If Err <> 0 Then
        MsgBox "保险帐户表的结构不正确，需要增加字段“就诊序号”！", vbInformation, gstrSysName
        Exit Function
    End If
    
    int序号 = Val(Nvl(rsTemp!就诊序号, 0))
    If bln更新 Then
        int序号 = int序号 + 1
        gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'就诊序号','''" & int序号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
    End If
    Get就诊序号 = int序号
End Function

Private Function GetMode(ByVal lng病人ID As Long, ByVal lng主页ID As Long, str住院门诊号 As String) As Boolean
    'MODIFIED BY ZYB 20040626 '三院修改，由于三晶老系统是按病人ID上传的，因此，流水号改为：
    '如果保险帐户中存在住院门诊号，且值不为空，则以它做为病人标识上传；否则按现有模式上传
    Dim bln模式 As Boolean              '为真，按病人ID_主页ID_就诊序号返回；否则仅返回病人ID
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    gstrSQL = " Select 住院门诊号 From 保险帐户 " & _
              " Where 险类=[1] ANd 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取住院门诊号", TYPE_重庆市, lng病人ID)
    If Err <> 0 Then
        '说明不存在该字段
        bln模式 = True
    Else
        bln模式 = (Nvl(rsTemp!住院门诊号) = "")
        If Not bln模式 Then str住院门诊号 = Nvl(rsTemp!住院门诊号)
    End If
    GetMode = bln模式
End Function
'
'Private Function GetRegisted(ByVal str医保号 As String) As Long
'    Dim strDate As String, strStart As String, strEnd As String
'    Dim rsTemp As New ADODB.Recordset
'
'    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
'    strStart = strDate & " 00:00:00"
'    strEnd = strDate & " 23:59:59"
'    '如果当天内存在就诊记录(挂号或收费)，则返回病人ID，否则返回零
'    gstrSQL = " Select A.病人ID From 病人费用记录 A,保险结算记录 B" & _
'              " Where A.记录性质 In (1,4) And A.结帐ID Is Not NULL" & _
'              " And A.登记时间 Between to_date('" & strStart & "','yyyy-MM-dd hh24:mi:ss')" & _
'              " And to_date('" & strEnd & "','yyyy-MM-dd hh24:mi:ss')" & _
'              " And A.病人ID+0 =(Select 病人ID From 保险帐户 Where 险类=" & TYPE_重庆市 & " ANd 医保号='" & str医保号 & "')" & _
'              " And A.结帐ID=B.记录ID And B.性质=1"
'    Call OpenRecordset(rsTemp, "取结帐ID")
'    If rsTemp.RecordCount = 0 Then Exit Function
'    GetRegisted = rsTemp!病人ID
'End Function

Private Function GetRegisted(ByVal str医保号 As String, ByRef str就诊时间 As String) As Long
    '如果当天存在就诊登记记录则说明就诊过
    Dim strDate As String, strStart As String, strEnd As String
    Dim rsTemp As New ADODB.Recordset

    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    strStart = strDate & " 00:00:00"
    strEnd = strDate & " 23:59:59"
    
    '一天虽产生多条就诊登记记录，但其医保流水号是不变的，因此随便取一条
    gstrSQL = " Select A.病人ID,A.就诊时间" & _
              " From 就诊登记记录 A,保险帐户 B" & _
              " Where A.主页ID=0 And A.记录ID Is Not NULL And A.病人ID=B.病人ID And A.险类=B.险类 " & _
              " And B.险类=[1] And B.医保号=[2]" & _
              " And A.就诊时间 Between [3] And [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断该病人当天是否就诊过", TYPE_重庆市, str医保号, CDate(strStart), CDate(strEnd))
    If rsTemp.RecordCount = 0 Then Exit Function
    
    str就诊时间 = Format(rsTemp!就诊时间, "yyyy-MM-dd HH:mm:ss")
    GetRegisted = rsTemp!病人ID
End Function

Public Function 取消就诊_重庆市(ByVal bytType As Byte, ByVal lng病人ID As Long)
    gstrSQL = "zl_就诊登记记录_DELETE(" & TYPE_重庆市 & "," & lng病人ID & ",0,to_date('" & gstr就诊时间 & "','yyyy-MM-dd hh24:mi:ss'))"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
End Function

Private Function MessageInfo(ByVal strInfo As String, ByVal bln批量预结算 As Boolean) As String
    '如果提示，则MSGBOX，返回值为空；否则不提示，返回值等于将要显示的信息
    If Not bln批量预结算 Then
        MsgBox strInfo, vbInformation, gstrSysName
    Else
        MessageInfo = strInfo
    End If
End Function

Private Function Is家庭病床(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 出院病床 From 病案主页 Where 当前病区ID is not null And 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否家庭病床", lng病人ID, lng主页ID)
    Is家庭病床 = IsNull(rsTemp!出院病床)
End Function
