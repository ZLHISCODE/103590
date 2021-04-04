Attribute VB_Name = "mdl贵阳"
Option Explicit
#Const gblnTest = 0     '1-调试

'总体修改说明：
'修改时间：2005-01-14
'修改人：朱玉宝
'修改内容：增加两个接口(SetBearingFlag、UploadICD)，其他大部分接口增加了入参与出参

'修改时间：2008-07-03
'修改人：朱玉宝
'修改内容：增加居民，工伤保险
'详细清单如下
'1 ?增加——获取工伤认定函数: GETGSINFO ok
'2 ?增加——查询中心特殊病种费用明细数据: QUERYCENSPECFEELIST   以前查询都没做，本次也不做
'3、增加——工伤保险清算申请与撤销：APPRECG，DELRECG    ok
'4、居民只能做特殊门诊与住院，工伤只能做普通门诊与住院，HIS不需要限制吗？不需要 ok
'5、获取医院单病种清算数据函数：QUERYHOSPSINGLEILLNESS，为何只支持职工医保与居民医保 ok
'6、查询医院单病种包干结算目录：QUERYHOSPSINGLEILLNESS_BG，同上     ok
'7、增加——查询中心特殊病种费用明细数据：QUERYCENSPECFEELIST   与2重复了
'8 ?普通门诊结算增加入参: 工伤认定编号 ok
'9、居民医保使用特殊门诊结算？处方本编号是居民特有的吗？  ok
'10、入院登记支持工伤及居民医保，取消入参：参保前已在院  ok
'11 ?住院特殊结算取消入院: 转大额标志 没找到此参数
'12、医疗保险申报清算：APPRECM 入参的参数名改名太大，适用于居民 ok
'13 ?撤销医疗保险清算: DELRECM 增加入参: 保险类别   ok
'14 ?撤销生育保险清算大动   无变化呢,估计当前对比错误
'15、人员类别    11：在职；21：退休；32：省属离休；34：市属离休；41：普通居民；42：低保对象；43：三无人员；44：低收入家庭；45：重度残疾； ok
'16、医院类别    01：一级；02：二级；03：三级；04：乡镇卫生所；05：药店；06：社区医院；09：非定点
'17、保险类别    1:企业基本医疗保险；2:企业离休医疗保险；3:机关事业单位基本医疗保险；4：企业生育保险；5：机关事业单位生育保险；6：居民医保；7：工伤保险  ok
'18、项目支付类别    11-普通门诊 12-普通住院 21-离休门诊 22-离休住院 31-生育门诊 32-生育住院 41-工伤门诊 42-工伤住院

Public mdomInput As MSXML2.DOMDocument
Public mdomOutput As MSXML2.DOMDocument
Public tdomInput As MSXML2.DOMDocument
'========================================================================================
'=日志类型
'========================================================================================
Public Enum LogType
    DBConnLTNew = 0
    DBConnLTEdit = 1
    DBConnLTDelete = 2
    DBConnLTSping = 3
End Enum

Public gblnHIS1026 As Boolean 'HIS版本是否是10.26或以上版本
Public gblnHIS1029 As Boolean 'HIS版本是否是10.29或以上版本【需10.29.40】
Public mlngCloseTime As Long '门诊结算窗口自动关闭时间
Public gbln慢病用药限制 As Boolean, gint累计用药量计算标准 As Integer, ging病历检查 As Integer
Public gint进行医保交易 As Integer
Private mblnInit As Boolean
Private mblnFail As Boolean
Public mstr卡号 As String
Public mstr密码 As String
Public mstr医保号 As String
Private mdbl余额 As Double
Private mlng病人ID As Long
Private mbln医保出院 As Boolean         '出院时是否同步办理医保出院
Private mbln特殊药品提示 As Boolean
Private mbln无密码结算 As Boolean
Public gstr工伤认定编号 As String
Public gint工伤康复住院 As Integer      '0-否;1-是

Public gintType As Integer              '卡类型
Public gstrSNO As String                '社保卡号
Public gstrIDNO As String               '身份证号
Public gstrPSAMNO As String             'PSAM芯片编号
Public gstrClientIP As String           '客户端IP

Public dblTOTAL As Double
'仅用于工伤zyq20110812增加
Public gstr工伤信息 As String
Public gstr门诊标志 As String
Public gstr工伤标志 As String
'仅用于门诊
Private mint结算方式 As Integer
Private mstr单病种编码 As String
Public gstr处方本号 As String

Public gobj贵阳 As Object
Private obj清算 As Object
Public Const mstr医保中心编码_贵阳 As String = "0101"
Public gcnGYYB As New ADODB.Connection

Private Type balance
    dbl医保基金 As Double
    dbl个人帐户 As Double
    dbl大病统筹 As Double
    dbl医疗补助 As Double
    dbl差额记帐 As Double
End Type

'多单据收费时提示用
Public Type t门诊数据
    dbl医保基金 As Double
    dbl大病基金 As Double
    dbl公务员补助 As Double
    dbl个人帐户 As Double
    dbl现金 As Double
    dbl费用总额 As Double
    dbl差额记帐 As Double
End Type
Public g门诊数据 As t门诊数据

Public gbln批量虚拟结算 As Boolean      '如果为true,不弹出msgbox提示,否则提示,主要是用于批量虚拟结算
Private preBalance As balance

Private mnodRowset As MSXML2.IXMLDOMElement

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long) As Long

'--读卡相关函数
Public Declare Function SGZ_IFD_Open Lib "SGZ_SSSReader.dll" (ByVal iReaderPort As Long, ByRef iReaderHandle As Long, ByVal iERRInfo As String) As Long
Public Declare Function SGZ_SAM_ReadNmuber Lib "SGZ_SSSReader.dll" (ByVal iReaderHandle As Long, ByVal iOutFileData As String, ByVal iERRInfo As String) As Long
Public Declare Function SGZ_ICC_ReadCardInfo Lib "SGZ_SSSReader.dll" (ByVal iReaderHandle As Long, ByVal iCardType As Long, ByVal iPassword As String, ByVal iInputFileAddr As String, ByVal iOutFileData As String, ByVal iERRInfo As String) As Long
Public Declare Function SGZ_IFD_GetPIN Lib "SGZ_SSSReader.dll" (ByVal iReaderHandle As Long, ByVal iDevType As Long, ByVal szPasswd As String, ByVal iERRInfo As String) As Long
Public Declare Function SGZ_IFD_Close Lib "SGZ_SSSReader.dll" (ByVal iReaderHandle As Long, ByVal iERRInfo As String) As Long

'启用补充结算
'========================================================================================
'==针对中心结算成功，本地未成功的结算记录进行补充
'==XieRong 2010-10-12
'========================================================================================
Public Type typ补充结算

    blnYn                       As Boolean
    str卡号                     As String
    str医保号                   As String
    lng病人ID                   As Long
    lng主页ID                   As Long
    STR姓名                     As String
    str住院号                   As String
    m_全自付                    As Double
    m_挂钩自付                  As Double
    m_起付线                    As Double
    m_基数自付                  As Double
    m_统筹支付                  As Double
    m_统筹自付                  As Double
    m_大病统筹                  As Double
    m_大病自付                  As Double
    m_超限自付                  As Double
    m_医保总费用                As Double
    m_公务员补助                As Double
    m_结算总费用                As Double
    m_HIS总费用                 As Double
    m_结算编号                  As String
    m_就诊顺序号                As String
    m_结算日期                  As String
    m_公务员补助起付标准        As Double
    m_公务员补助起付线          As Double
    m_普通门诊公务员补助累计    As Double
    m_超大额限额公务员补助      As Double
    m_特殊结算方式              As String
    m_特殊结算说明              As String

End Type
Public g补充结算        As typ补充结算
Private mlng结帐ID              As Long


Public Function 医保初始化_贵阳() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    Dim strUser As String, strPass As String, strServer As String
    Dim rsTemp As New ADODB.Recordset
    
    If mblnInit Then
        医保初始化_贵阳 = True
        Exit Function
    End If
    If mblnFail Then Exit Function
    
    On Error Resume Next
    If gstrClientIP = "" Then gstrClientIP = zl_Ip_Address_FromOrc(gstrClientIP)
    
    gint进行医保交易 = GetSetting(appName:="ZLSOFT", Section:="公共全局", Key:="贵阳医保交易", Default:="1")
    If gint进行医保交易 = 1 Then
        '需要连接医保时
        Set mdomInput = New MSXML2.DOMDocument
        If Err <> 0 Then
            ShowMsgbox "不能创建XML分析器，请注册msxml3.dll部件。"
            Exit Function
        End If
        
        Dim strYBServer As String
        On Error Resume Next
        #If gblnTest = 1 Then
            Set gobj贵阳 = CreateObject("GYSYB.CLSGYSYB")
            If Err <> 0 Then
                ShowMsgbox "加载调试部件时出错，错误信息如下：" & vbCrLf & Err.Description
                Exit Function
            End If
            Set obj清算 = gobj贵阳
        #Else
            '如果用全局变量，有时调用时会等很久，可能资源分配的原故
            strYBServer = Get保险参数_贵阳("医保服务器")
            If strYBServer = "" Then
                Set gobj贵阳 = CreateObject("HospCOMSvr.HospCOMServer")
                Set obj清算 = CreateObject("HospRecSvr.HospRecServer")
            Else
                Set gobj贵阳 = CreateObject("HospCOMSvr.HospCOMServer", strYBServer)
                Set obj清算 = CreateObject("HospRecSvr.HospRecServer", strYBServer)
            End If
            If Err <> 0 Then
                mblnFail = True
                ShowMsgbox "无法创建医保接口部件（HospCOMSvr.HospCOMServer）。"
                Exit Function
            End If
        #End If
    End If
    '取保险参数
    On Error GoTo errHand
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & TYPE_贵阳市
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险参数", TYPE_贵阳市)
    Do While Not rsTemp.EOF
        If rsTemp!参数名 = "医保用户名" Then
            strUser = Nvl(rsTemp!参数值)
        ElseIf rsTemp!参数名 = "医保用户密码" Then
            strPass = Nvl(rsTemp!参数值)
        ElseIf rsTemp!参数名 = "医保服务器1" Then
            strServer = Nvl(rsTemp!参数值)
        ElseIf rsTemp!参数名 = "出院操作" Then
            mbln医保出院 = (Nvl(rsTemp!参数值, 0) = 0)
        ElseIf rsTemp!参数名 = "特殊药品提示" Then
            mbln特殊药品提示 = (Nvl(rsTemp!参数值, 0) = 1)
        ElseIf rsTemp!参数名 = "门诊结算窗口关闭时间" Then
            mlngCloseTime = Nvl(rsTemp!参数值, 20)
        ElseIf rsTemp!参数名 = "启用门诊慢性疾病用药限制功能" Then
            gbln慢病用药限制 = (Val(Nvl(rsTemp!参数值, 0)) = 1)
        ElseIf rsTemp!参数名 = "累计用药量计算标准" Then
            gint累计用药量计算标准 = Nvl(rsTemp!参数值, 2)
        ElseIf rsTemp!参数名 = "病历检查" Then
            ging病历检查 = Nvl(rsTemp!参数值, 0)
        End If
        rsTemp.MoveNext
    Loop
    If Not OraDataOpen(gcnGYYB, strServer, strUser, strPass, True) Then Exit Function
    
    'HIS版本号 程池富，10.26以上及以下版本兼容 2011-04-07
    gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
    If Split(rsTemp!版本号, ".")(0) = 10 And Split(rsTemp!版本号, ".")(1) >= 29 Then
        '10.29以上版本
        gblnHIS1029 = True
    ElseIf Split(rsTemp!版本号, ".")(0) = 10 And Split(rsTemp!版本号, ".")(1) >= 26 Then
        '10.26及以上高版本
        gblnHIS1026 = True
        gblnHIS1029 = False
    Else
        '10.26及以下低版本
        gblnHIS1026 = False
        gblnHIS1029 = False
    End If
         
    mblnInit = True
    医保初始化_贵阳 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 身份标识_贵阳(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim str保险类别 As String, str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, cur帐户余额 As Currency
    Dim STR姓名 As String, str性别 As String, str身份证号码 As String, lng年龄 As Long
    Dim str出生日期 As String, str人员类别 As String, str单位编码 As String, str单位名称 As String, str医疗照顾人员 As String
    Dim strIdentify As String, str附加 As String
    Dim bln生育标志 As Boolean, str支付类型 As String
    Dim lng病种ID As Long
    Dim rsTemp As New ADODB.Recordset
    
    '初始化一些变量，在程序中途退出时值却已经赋了
    gstr工伤认定编号 = ""
    If frmIdentify贵阳.GetIdentify(bytType, str卡号, str医保号, str分中心编号, str密码, bln生育标志, lng病种ID, str支付类型) = False Then
        Exit Function
    End If
    If bytType = 0 Or bytType = 3 Then str支付类型 = ""
    '还原数据
    str保险类别 = Split(str卡号, "^")(1)
    str卡号 = Split(str卡号, "^")(0)
    
    If bytType = id门诊确认 Then
        '该返回值暂时没有作用，只要不为空就表示成功了
        身份标识_贵阳 = str卡号 & ";" & str医保号 & ";" & str密码
        Exit Function
    End If
    
    '取得返回值
    STR姓名 = GetElemnetValue("PERSONNAME")
    str性别 = GetElemnetValue("SEX")
    str性别 = Switch(str性别 = "1", "男", str性别 = "2", "女", str性别 = "9", "其它", True, str性别)
    str身份证号码 = GetElemnetValue("PID")
    
    str出生日期 = AddDate(GetElemnetValue("BIRTHDAY"))
    If IsDate(str出生日期) = True Then
        lng年龄 = DateDiff("yyyy", CDate(str出生日期), zlDatabase.Currentdate)
    Else
        str出生日期 = ""
    End If
    
    str人员类别 = GetElemnetValue("PERSONTYPE")
    str人员类别 = Switch(str人员类别 = "11", "在职", str人员类别 = "21", "退休", _
                      str人员类别 = "32", "省属离休", str人员类别 = "34", "市属离休", _
                      str人员类别 = "41", "普通居民", str人员类别 = "42", "低保对象", _
                      str人员类别 = "43", "三无人员", str人员类别 = "44", "低收家庭", _
                      str人员类别 = "45", "重度残疾", True, "其他") '界面上仍显示低收入家庭，但数据库只有8位，只能保存为低收家庭
    str单位编码 = ToVarchar(GetElemnetValue("DEPTCODE"), 12)
    str单位名称 = ToVarchar(GetElemnetValue("DEPTNAME"), 36) '字段长度本是50，但由于还要保存编码及括号
    cur帐户余额 = Val(GetElemnetValue("ACCTBALANCE"))
    str保险类别 = Val(GetElemnetValue("INSURETYPE"))
    str医疗照顾人员 = Val(GetElemnetValue("CAREPSNFLAG"))
    
    '卡号;医保号;密码;姓名;性别;出生日期;身份证;工作单位
    '医保号第一位为卡类型
    '把分号替换成逗号
    strIdentify = Replace(str卡号, ";", ",") & ";" & str医保号 & ";;" & STR姓名 & ";" & str性别 & ";" & str出生日期 & ";" & str身份证号码 & ";" & str单位名称 & "(" & str单位编码 & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    str附加 = ";"                                       '8.中心代码
    str附加 = str附加 & ";"                             '9.顺序号  但本医保用于保存医保分中心编码（避免建立医保中心）
    str附加 = str附加 & ";" & str人员类别               '10人员身份
    str附加 = str附加 & ";" & cur帐户余额               '11帐户余额
    str附加 = str附加 & ";0"                            '12当前状态
    str附加 = str附加 & ";" & IIf(lng病种ID <> 0, lng病种ID, "")   '13病种ID
    str附加 = str附加 & ";" & IIf(str人员类别 = "在职", 1, 2)      '14在职(1,2)
    str附加 = str附加 & ";"                             '15退休证号
    str附加 = str附加 & ";" & lng年龄                   '16年龄段
    str附加 = str附加 & ";"                             '17灰度级
    str附加 = str附加 & ";" & cur帐户余额               '18帐户增加累计
    str附加 = str附加 & ";0"                            '19帐户支出累计
    str附加 = str附加 & ";"                             '20进入统筹累计
    str附加 = str附加 & ";"                             '21统筹报销累计
    str附加 = str附加 & ";"                             '22住院次数累计
    str附加 = str附加 & ";"                             '23就诊类型 (1、急诊门诊)
    
    lng病人ID = BuildPatiInfo(bytType, strIdentify & str附加, lng病人ID, TYPE_贵阳市)
    '返回格式:中间插入病人ID
    If lng病人ID <> 0 Then
        身份标识_贵阳 = strIdentify & ";" & lng病人ID & str附加
        
        mstr卡号 = str卡号
        mstr密码 = str密码
        
        '更新当前医保病人的保险类别以及医疗照顾人员标志
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'保险类别','''" & str保险类别 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险类别")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'医疗照顾人员','''" & str医疗照顾人员 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存医疗照顾人员")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'生育标志','''" & IIf(bln生育标志, 1, 0) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存生育标志")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'支付类型','''" & str支付类型 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存支付类型")
    
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'卡类别','''" & gintType & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存卡类别")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'IDNO','''" & gstrIDNO & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存身份证号")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'PSAM','''" & gstrPSAMNO & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存PSAM")
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 更改密码_贵阳市(ByVal str磁卡数据 As String, ByVal str密码 As String, ByVal str新密码 As String) As Boolean
    If InitXML = False Then Exit Function
    
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' 卡类别
    Call InsertChild(mdomInput.documentElement, "CARDDATA", str磁卡数据)            ' 磁卡数据（或者IC卡号）
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' 社会保障号（身份证号）
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP地址
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM卡芯片
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str密码)                ' 密码
    Call InsertChild(mdomInput.documentElement, "NEWPASSWORD", str新密码)           ' 密码
    
    '调用接口
    If CommServer("MODIFYCARD") = False Then Exit Function
    更改密码_贵阳市 = True
End Function

Public Function 个人余额_贵阳(strSelfNo As String) As Currency
'功能: 提取参保病人个人帐户余额
'参数: strSelfNO-病人个人编号
'返回: 返回个人帐户余额的金额
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    '从数据库中读取（因为刚才才保存了的，应该是准确的）
    gstrSQL = " Select /*+ rule */  B.帐户余额 " & _
              " From 医保病人关联表 A,医保病人档案 B " & _
              " Where A.险类=B.险类 AND A.中心=B.中心 AND A.医保号=B.医保号 AND A.标志=1 " & _
              " And B.险类=[1] and B.中心=0 and B.医保号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "贵阳医保", TYPE_贵阳市, strSelfNo)
    
    If rsTemp.EOF = False Then
        个人余额_贵阳 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_贵阳(rs明细 As ADODB.Recordset, str结算方式 As String, Optional strAdvance As String) As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim str项目编码 As String
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, str人员类别 As String, str保险类别 As String
    Dim dbl个人帐户 As Double, dbl公务员补助 As Double, dblHIS总额 As Double, dbl结算总费用 As Double, dbl医保基金 As Double, dbl大病统筹 As Double, dbl差额 As Double
    Dim lng病人ID As Long, str疾病编码 As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Dim int单据_CUR As Integer, int单据_MAX As Integer, str药品超量 As String
    Dim str收费细目IDS      As String   '收费细目编码

    On Error GoTo errHandle
    
    '增加单病种以后，返回参数增加：结算总费用，我们需要增加个差额字段，用来保证HIS计算出来的现金支付与医保一致，公式如下：
    '差额=HIS总费用-结算总费用；现金支付=HIS总费用-统筹支付-差额
    
    If rs明细.RecordCount = 0 Then
        str结算方式 = "个人帐户;0;0"
        门诊虚拟结算_贵阳 = True
        Exit Function
    End If
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID")
    
    '判断该病人是否是特殊门诊
    gstrSQL = "select A.人员身份,A.保险类别,B.编码 from 保险帐户 A,保险病种 B where A.病人ID=[1] and A.险类=[2] and A.病种ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", lng病人ID, TYPE_贵阳市)
    If rsTemp.EOF = False Then
        str疾病编码 = IIf(IsNull(rsTemp("编码")), "", rsTemp("编码"))
        str人员类别 = Nvl(rsTemp("人员身份"), "")
        str保险类别 = Nvl(rsTemp!保险类别)
        '转换人员身份
        str人员类别 = Switch(str人员类别 = "在职", "11", str人员类别 = "退休", "21", _
                          str人员类别 = "省属离休", "32", str人员类别 = "市属离休", "34", _
                          str人员类别 = "普通居民", "41", str人员类别 = "低保对象", "42", _
                          str人员类别 = "三无人员", "43", str人员类别 = "低收家庭", "44", _
                          str人员类别 = "重度残疾", "45", True, "11") '界面上仍显示低收入家庭，但数据库只有8位，只能保存为低收家庭
    End If
    datCurr = zlDatabase.Currentdate
    '分解单据总数，当前单据
    If strAdvance = "" Then strAdvance = "1|1"
    int单据_CUR = Val(Split(strAdvance, "|")(1))
    int单据_MAX = Val(Split(strAdvance, "|")(0))
'    If (int单据_MAX > 1) And 多单据收费_收费分别打印 Then
'        MsgBox "请取消系统参数设置中票据页面里的参数“门诊收费每张单据分别打印”，方可进行多单据收费！", vbInformation, gstrSysName
'        Exit Function
'    End If
    mint结算方式 = 0: mstr单病种编码 = ""
    If str疾病编码 = "" Then
ReChoose:
        '普通门诊要求选择结算方式与单病种结算目录（结算方式;单病种编码）
        If int单据_CUR = 1 Then
            mstr单病种编码 = 设置结算方式_贵阳(lng病人ID, Nothing, False)
            If mstr单病种编码 = "" Then mstr单病种编码 = ";"
            mint结算方式 = Val(Split(mstr单病种编码, ";")(0))
            mstr单病种编码 = Split(mstr单病种编码, ";")(1)
        End If
    End If
    
    'If Get验证_贵阳(0, str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
    
    
    
    If int单据_CUR = 1 Then
        dblTOTAL = 0
        preBalance.dbl差额记帐 = 0
        preBalance.dbl大病统筹 = 0
        preBalance.dbl个人帐户 = 0
        preBalance.dbl医保基金 = 0
        preBalance.dbl医疗补助 = 0
        
        '对XML DomDocument对象进行初始化
        If InitXML = False Then Exit Function
        Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' 卡类别
        Call InsertChild(mdomInput.documentElement, "CARDDATA", mstr卡号)           ' 磁卡编码
        Call InsertChild(mdomInput.documentElement, "INSURETYPE", str保险类别)     ' 保险类别
        Call InsertChild(mdomInput.documentElement, "PASSWORD", mstr密码)         ' 密码
        Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstr医保号)      ' 个人编号
        Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' 社会保障号
        Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP地址
        Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM卡芯片
        If str疾病编码 <> "" Then '特殊门诊
            '补足8位长度
            str疾病编码 = String(8 - Len(str疾病编码), "0") & str疾病编码
            Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str疾病编码)         '特种病编码
            Call InsertChild(mdomInput.documentElement, "CFNO", gstr处方本号)           '处方本号
        End If
        If str保险类别 = "7" Then
            Call InsertChild(mdomInput.documentElement, "GSRDBH", gstr工伤认定编号)          '工伤认定编号
        End If
        Call InsertChild(mdomInput.documentElement, "ISCAL", 0)         ' 是否结算
        Call InsertChild(mdomInput.documentElement, "CALTYPE", mint结算方式)         ' 结算方式
        Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", mstr单病种编码)         ' 单病种结算编码
        Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", "0")     ' 账户支付额
        Call InsertChild(mdomInput.documentElement, "INVOICENO", "") ' 发票号
        Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' 操作员
        Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) ' 办理日期
        Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) '待遇开始享受时间
        Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' 消费明细
        Set mnodRowset = nodRowset
    Else
        Set nodRowset = mnodRowset
        Set mdomInput = tdomInput
    End If
    str收费细目IDS = ""
    Do Until rs明细.EOF
        If Nvl(rs明细!实收金额, 0) <> 0 Then
            gstrSQL = "SELECT C.药品ID,C.规格,E.名称 AS 剂型  FROM 药品目录 C,药品信息 D,药品剂型 E WHERE C.药品ID=" & rs明细("收费细目ID") & " AND C.药名ID=D.药名ID AND D.剂型=E.编码"
            gstrSQL = "select A.类别,A.名称,B.项目编码,nvl(A.规格,F.规格) AS 规格,F.剂型,A.计算单位 from 收费细目 A,保险支付项目 B,(" & gstrSQL & _
                    ") F where A.ID=[1] and A.ID=B.收费细目ID  AND A.Id=F.药品ID(+) and B.险类=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", CLng(rs明细("收费细目ID")), TYPE_贵阳市)
            If rsTemp.EOF = True Then
                MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
                Exit Function
            End If
             '药品使用限量检测 2011-05-21  程池富
            If gbln慢病用药限制 = True Then
              str药品超量 = str药品超量 & DrugsUsed(TYPE_贵阳市, lng病人ID, Nvl(rs明细!收费细目ID), Nvl(rs明细!数量, 0))
            End If
        

            Set nodRow = InsertChild(nodRowset, "ROW", "")
            
            On Error Resume Next
            str项目编码 = ""
            str项目编码 = Nvl(rs明细!保险编码)
            Err = 0
            On Error GoTo errHandle
            Call 贵阳特殊药品提示(rs明细!收费细目ID)
                    '对收费细目ID判断是否处于黑名单
        
            str收费细目IDS = str收费细目IDS & "," & Nvl(rs明细!收费细目ID)

            '如果保险编码为空，则需要用户选择编码
            If str项目编码 = "" Then
                str项目编码 = GetItemInsure_贵阳(lng病人ID, rs明细!收费细目ID, True)
            End If
            If str项目编码 = "" Then str项目编码 = Nvl(rsTemp!项目编码)
            
            Call nodRow.setAttribute("ITEMCODE", ToVarchar(str项目编码, 12))
            Call nodRow.setAttribute("ITEMNAME", ToVarchar(rsTemp("名称"), 72))
            Call nodRow.setAttribute("SUBJECT", Subject(rsTemp("类别")))
            Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rsTemp("规格"), 40))
            Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rsTemp("剂型"), 20))
            Call nodRow.setAttribute("UNIT", ToVarchar(rsTemp("计算单位"), 20))
            Call nodRow.setAttribute("PRICE", Format(rs明细("单价"), "0.0000"))
            Call nodRow.setAttribute("QUANTITY", Format(rs明细("数量"), "0.00"))
            Call nodRow.setAttribute("FROMOFFICE", ToVarchar(UserInfo.部门, 56)) '开单科室
            Call nodRow.setAttribute("FROMDOCT", Format(UserInfo.姓名, 20))      '开单医生
            Call nodRow.setAttribute("TOOFFICE", ToVarchar(UserInfo.部门, 56))  '受单科室
            Call nodRow.setAttribute("TODOCT", Format(UserInfo.姓名, 20))       '受单医生
            Call nodRow.setAttribute("DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))        '办理日期
            Call nodRow.setAttribute("NOTE", ToVarchar(rs明细("摘要"), 512))        '备注
            
            dblTOTAL = dblTOTAL + Round(rs明细!实收金额, 2)
        End If
        rs明细.MoveNext
    Loop
    If Len(Replace(str药品超量, vbNewLine, "")) <> 0 Then ShowMsgbox "第" & int单据_CUR & "张收费单据中：" & vbNewLine & str药品超量: Exit Function
    
    gstrSQL = "select distinct 医保号 FROM  保险帐户  WHERE 险类=[1] AND 病人ID=[2]"
    str医保号 = zlDatabase.OpenSQLRecord(gstrSQL, "医保黑名单", TYPE_贵阳市, lng病人ID).Fields(0)
    '检测收费细目IDS是否存在于黑名单表
    str收费细目IDS = Mid(str收费细目IDS, 2)
    '检测当前医保号是否牌黑名单且启用
    gstrSQL = "select Count(1) from 医保黑名单_贵阳 where 医保号 = [1] and 状态 = 1"
    If Val(zlDatabase.OpenSQLRecord(gstrSQL, "医保黑名单", str医保号).Fields(0)) = 1 Then
        '启用医保号状态下检测是否有项目处于黑名单内
        gstrSQL = "Select A.医保号, A.收费细目id, C.编码, C.名称, C.规格" & vbCrLf & _
                 "From 医保黑名单项目_贵阳 A, Table(F_Str2list2([2])) B, 收费细目 C" & vbCrLf & _
                 "Where rownum = 1 And a.收费细目ID = B.C2 And a.收费细目ID = C.ID And a.医保号 = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保黑名单项目", str医保号, str收费细目IDS)
        '周玉强增加只限制特殊门诊和工伤保险病人
        If Not (rsTemp.EOF Or rsTemp.BOF) And (gstr门诊标志 = "特殊门诊" Or gstr工伤标志 = "工伤保险") Then
            '存在黑名单项目
            MsgBox "当前医保号【" & str医保号 & "】" & vbCrLf & _
                   "收费细目ID【" & rsTemp!收费细目ID & "】" & vbCrLf & _
                   "收费细目编码【" & rsTemp!编码 & "】" & vbCrLf & _
                   "收费细目名称【" & rsTemp!名称 & "】" & vbCrLf & _
                   "收费细目规格【" & rsTemp!规格 & "】" & vbCrLf & _
                   "存在于医保黑名单项目中！请医保相关人员先取消或禁用！", vbCritical, gstrSysName
            Exit Function
        End If
    End If
    
    '调用接口
    Set tdomInput = mdomInput
    If CommServer(IIf(str疾病编码 <> "", "CALSPECCLIN", "CALCLIN")) = False Then Exit Function
    
    '不同的人群，返回的XML的字段表示含义不同，不能直接取，所以要分别判断
    '离休人员不存在普通门诊与特殊门诊，统一由ALLOWFUND支付；
    '基本医疗人员特殊门诊由FUND1PAY及FUND2PAY支付，普通门诊由个人帐户支付
    If str人员类别 = "32" Or str人员类别 = "34" Then
        dbl医保基金 = Val(GetElemnetValue("ALLOWFUND"))
        dbl医保基金 = dbl医保基金 - preBalance.dbl医保基金
        preBalance.dbl医保基金 = Val(GetElemnetValue("ALLOWFUND"))
        str结算方式 = "医保基金;" & dbl医保基金 & ";0"
    Else
        dbl个人帐户 = Val(GetElemnetValue("ACCTPAY"))
        dbl个人帐户 = dbl个人帐户 - preBalance.dbl个人帐户
         preBalance.dbl个人帐户 = Val(GetElemnetValue("ACCTPAY"))
        str结算方式 = "个人帐户;" & dbl个人帐户 & ";1"   '允许修改个人帐户
        '以前是特殊门诊才返回，现在改为直接取，就值就报销，因为增加了工伤 Modified by ZYB 20080702
        
'        If str疾病编码 <> "" Then
            
            dbl医保基金 = Val(GetElemnetValue("FUND1PAY"))
            dbl医保基金 = dbl医保基金 - preBalance.dbl医保基金
            preBalance.dbl医保基金 = Val(GetElemnetValue("FUND1PAY"))
            
            dbl大病统筹 = Val(GetElemnetValue("FUND2PAY"))
            dbl大病统筹 = dbl大病统筹 - preBalance.dbl大病统筹
             preBalance.dbl大病统筹 = Val(GetElemnetValue("FUND2PAY"))
            
            str结算方式 = str结算方式 & "|医保基金;" & dbl医保基金 & ";0" & _
                         "|大病统筹;" & dbl大病统筹 & ";0"
'        End If
    End If
    dbl公务员补助 = Val(GetElemnetValue("FUND3PAY"))
    dbl公务员补助 = dbl公务员补助 - preBalance.dbl医疗补助
    preBalance.dbl医疗补助 = Val(GetElemnetValue("FUND3PAY"))
    str结算方式 = str结算方式 & "|医疗补助;" & dbl公务员补助 & ";0"
    
    dbl结算总费用 = Val(Format(GetElemnetValue("CALFEEALL"), "#0.00;-#0.00;0;"))
    dblHIS总额 = Val(Format(GetElemnetValue("HOSPFEEALL"), "#0.00;-#0.00;0;"))
    
    '先比较总额是否一致
    If Format(dblTOTAL, "#0.00") <> Format(dblHIS总额, "#0.00") Then
        If Abs(Val(Format(dblTOTAL, "#0.00")) - Val(Format(dblHIS总额, "#0.00"))) > 0.5 Then
            MsgBox "HIS总额与医保接收到的总费用不一致，不允许结算！" & vbCrLf & _
                "HIS:" & Format(dblTOTAL, "#0.00") & Space(10) & "医保:" & Format(dblHIS总额, "#0.00"), vbInformation, gstrSysName
            Exit Function
        Else
            If MsgBox("HIS总额与医保接收到的总费用不一致，估计是单价精度不足引起的误差，是否结算？" & vbCrLf & _
                "HIS:" & Format(dblTOTAL, "#0.00") & Space(10) & "医保:" & Format(dblHIS总额, "#0.00"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    dbl差额 = dblHIS总额 - dbl结算总费用
    dbl差额 = dbl差额 - preBalance.dbl差额记帐
    preBalance.dbl差额记帐 = dblHIS总额 - dbl结算总费用
    If dbl差额 <> 0 And mint结算方式 = 1 Then
        '差额=HIS总费用-结算总费用；现金支付=HIS总费用-统筹支付-差额
        str结算方式 = str结算方式 & "|差额记帐;" & dbl差额 & ";0"
    End If
    门诊虚拟结算_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_贵阳(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, Optional ByVal bln挂号 As Boolean = False, Optional ByRef strAdvance As String = "") As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim str项目编码 As String
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, str人员类别 As String, str保险类别 As String, str特殊结算方式 As String
    Dim str医生 As String, str科室 As String, cur发生费用 As Double, cur医保总费用 As Double, datCurr As Date
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim lng病人ID  As Long, str疾病编码   As String, lng项目数 As Long, cur帐户余额 As Currency
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str结算方式     As String
    Dim intCurr As Integer, intMAX As Integer, dbl现金 As Double
    Dim blnOld As Boolean               '是否是老版HIS系统
    
    '---------------------------------------------------------------------------------------------
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
            
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, cur起付线 As Double
    Dim dblHIS总额 As Double, dbl结算总费用 As Double, dbl差额 As Double
    Dim cur公务员补助起付标准 As Double, cur公务员补助起付线 As Double, cur普通门诊公务员补助累计 As Double
    Dim cur公务员补助 As Double, cur超大额限额公务员补助 As Double
    Dim str就诊顺序号 As String, str结算编号 As String, str特殊结算说明 As String
    
    On Error GoTo errHandle
    
    cur帐户余额 = 个人余额_贵阳(strSelfNo)
    lng项目数 = Val(Get保险参数_贵阳("门诊最大项目数"))
    
     '多单据处理 2011-4-7 程池富
    #If gverControl < 2 Then
        blnOld = True
    #End If
    strAdvance = Decode(Trim(strAdvance), "", "1|1", strAdvance)
    intCurr = Split(strAdvance, "|")(1) '当前单据
    intMAX = Split(strAdvance, "|")(0) '最大单据
    If intCurr = 1 Then
        g门诊数据.dbl大病基金 = 0
        g门诊数据.dbl费用总额 = 0
        g门诊数据.dbl个人帐户 = 0
        g门诊数据.dbl公务员补助 = 0
        g门诊数据.dbl现金 = 0
        g门诊数据.dbl医保基金 = 0
        g门诊数据.dbl差额记帐 = 0
    End If
    
    gstrSQL = "SELECT Nvl(从属父号,Nvl(价格父号,序号)) AS 主序号 FROM " & IIf(gblnHIS1026 = True, "门诊费用记录", "病人费用记录") & _
             " WHERE 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0" & _
             " GROUP BY Nvl(从属父号,Nvl(价格父号,序号))"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "贵阳医保", lng结帐ID)
    If rs明细.RecordCount > lng项目数 Then
        Err.Raise 9000, gstrSysName, "门诊收费的项目数不能超过" & lng项目数 & "。"
        Exit Function
    End If
    
    gstrSQL = "Select A.ID,A.序号,A.收费细目ID,A.记录性质,A.记录状态,A.病人ID,A.NO,A.登记时间,A.开单人 as 医生,A.登记时间," & _
            "   A.数次*A.付数 as 数量,A.标准单价 as 实际价格,A.结帐金额," & _
            "   A.收费类别,A.保险编码,D.项目编码,B.名称 as 项目名称,C.名称 as 科室名称,nvl(B.规格,F.规格) AS 规格,F.剂型,B.计算单位,A.摘要 " & _
            " From (Select * From " & IIf(gblnHIS1026 = True, "门诊费用记录", "病人费用记录") & " Where 结帐ID=[1] And Nvl(实收金额,0)<>0 And Nvl(附加标志,0)<>9) A,收费细目 B,部门表 C,保险支付项目 D " & _
            "     ,(SELECT C.药品ID,C.规格,E.名称 AS 剂型  FROM " & IIf(gblnHIS1026 = True, "门诊费用记录", "病人费用记录") & " A,药品目录 C,药品信息 D,药品剂型 E WHERE A.结帐ID=[1] AND A.收费细目ID=C.药品ID AND C.药名ID=D.药名ID AND D.剂型=E.编码) F " & _
            " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID And A.收费细目ID=D.收费细目ID  AND A.ID=F.药品ID(+) And D.险类=[2] And Nvl(A.附加标志,0)<>9 And Nvl(A.记录状态,0)<>0" & _
            " Order by A.ID"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "贵阳医保", lng结帐ID, TYPE_贵阳市)
    If rs明细.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "没有填写收费记录"
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str医生 = ToVarchar(IIf(IsNull(rs明细("医生")), UserInfo.姓名, rs明细("医生")), 20)
    str科室 = ToVarchar(IIf(IsNull(rs明细("科室名称")), UserInfo.部门, rs明细("科室名称")), 56)
    datCurr = zlDatabase.Currentdate
    
    '一、费用明细传递
    
    '判断该病人是否是特殊门诊
    gstrSQL = "select A.人员身份,A.保险类别,B.编码 from 保险帐户 A,保险病种 B where A.病人ID=[1] and A.险类=[2] and A.病种ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", lng病人ID, TYPE_贵阳市)
    If rsTemp.EOF = False Then
        str疾病编码 = IIf(IsNull(rsTemp("编码")), "", rsTemp("编码"))
        str人员类别 = Nvl(rsTemp("人员身份"), "")
        str保险类别 = Nvl(rsTemp!保险类别)
        '转换人员身份
        str人员类别 = Switch(str人员类别 = "在职", "11", str人员类别 = "退休", "21", _
                          str人员类别 = "省属离休", "32", str人员类别 = "市属离休", "34", _
                          str人员类别 = "普通居民", "41", str人员类别 = "低保对象", "42", _
                          str人员类别 = "三无人员", "43", str人员类别 = "低收家庭", "44", _
                          str人员类别 = "重度残疾", "45", True, "11") '界面上仍显示低收入家庭，但数据库只有8位，只能保存为低收家庭
    End If
    
    'If Get验证_贵阳(0, str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
        
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' 卡类别
    Call InsertChild(mdomInput.documentElement, "CARDDATA", mstr卡号)           ' 磁卡编码
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", str保险类别)   ' 保险类别
    Call InsertChild(mdomInput.documentElement, "PASSWORD", mstr密码)         ' 密码
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstr医保号)      ' 个人编号
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' 社会保障号
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP地址
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM卡芯片
    If str疾病编码 <> "" Then '特殊门诊
        '补足8位长度
        str疾病编码 = String(8 - Len(str疾病编码), "0") & str疾病编码
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str疾病编码)         '特种病编码
        Call InsertChild(mdomInput.documentElement, "CFNO", gstr处方本号)          '处方本号
    End If
    If str保险类别 = "7" Then
        Call InsertChild(mdomInput.documentElement, "GSRDBH", gstr工伤认定编号)          '工伤认定编号
    End If
    Call InsertChild(mdomInput.documentElement, "ISCAL", 1)               ' 是否结算
    Call InsertChild(mdomInput.documentElement, "CALTYPE", mint结算方式)         ' 结算方式
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", mstr单病种编码)         ' 单病种结算编码
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", Format(cur个人帐户, "0.00")) ' 账户支付额
    Call InsertChild(mdomInput.documentElement, "INVOICENO", "M" & "_" & rs明细!记录性质 & "_" & rs明细("NO")) ' 发票号
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(rs明细("登记时间"), "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(rs明细("登记时间"), "yyyy-MM-dd HH:mm:ss")) '待遇开始享受时间
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' 消费明细
    
    dblHIS总额 = 0
    Do Until rs明细.EOF
        cur发生费用 = cur发生费用 + rs明细("结帐金额")

            
        Set nodRow = InsertChild(nodRowset, "ROW", "")
        
        str项目编码 = Nvl(rs明细!保险编码)
        
        '如果保险编码为空，则需要用户选择编码
        If str项目编码 = "" Then
            str项目编码 = GetItemInsure_贵阳(lng病人ID, rs明细!收费细目ID, False)
        End If
        If str项目编码 = "" Then str项目编码 = Nvl(rs明细!项目编码)
        
        Call nodRow.setAttribute("ITEMCODE", ToVarchar(str项目编码, 12))
        Call nodRow.setAttribute("ITEMNAME", ToVarchar(rs明细("项目名称"), 72))
        Call nodRow.setAttribute("SUBJECT", Subject(rs明细("收费类别")))
        Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rs明细("规格"), 40))
        Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rs明细("剂型"), 20))
        Call nodRow.setAttribute("UNIT", ToVarchar(rs明细("计算单位"), 20))
        Call nodRow.setAttribute("PRICE", Format(rs明细("实际价格"), "0.0000"))
        Call nodRow.setAttribute("QUANTITY", Format(rs明细("数量"), "0.00"))
        
        dblHIS总额 = dblHIS总额 + Format(rs明细("实际价格") * rs明细("数量"), "0.0000")
        
        Call nodRow.setAttribute("FROMOFFICE", str科室)    '开单科室
        Call nodRow.setAttribute("FROMDOCT", str医生)      '开单医生
        Call nodRow.setAttribute("TOOFFICE", str科室)     '受单科室
        Call nodRow.setAttribute("TODOCT", str医生)       '受单医生
        
        '处理时间时，为了保证同一保险项目的的收费时间不同，因此在登记时间上按序号加上秒数
        Call nodRow.setAttribute("DODATE", Format(rs明细("登记时间"), "yyyy-MM-dd HH:mm:ss"))    '办理日期
        Call nodRow.setAttribute("NOTE", ToVarchar(rs明细("摘要"), 512))         '备注
        
        rs明细.MoveNext
    Loop
     
    '保存结算记录
    '---------------------------------------------------------------------------------------------
    
    If g补充结算.blnYn And g补充结算.str卡号 = str卡号 And g补充结算.str医保号 = str医保号 And g补充结算.lng病人ID = lng病人ID Then
        g补充结算.blnYn = False
        cur全自付 = g补充结算.m_全自付
        cur挂钩自付 = g补充结算.m_挂钩自付
        cur起付线 = g补充结算.m_起付线
        cur基数自付 = g补充结算.m_基数自付
        
        If str人员类别 = "32" Or str人员类别 = "34" Then
            cur统筹支付 = g补充结算.m_统筹支付
            cur大病统筹 = 0
        Else
            cur统筹支付 = g补充结算.m_统筹支付
            cur大病统筹 = g补充结算.m_大病统筹
        End If
        cur统筹自付 = g补充结算.m_统筹自付
        cur大病自付 = g补充结算.m_大病自付
        cur超限自付 = g补充结算.m_超限自付
        cur公务员补助 = g补充结算.m_公务员补助
        
        cur公务员补助起付标准 = g补充结算.m_公务员补助起付标准
        cur公务员补助起付线 = g补充结算.m_公务员补助起付线
        cur普通门诊公务员补助累计 = g补充结算.m_普通门诊公务员补助累计
        cur超大额限额公务员补助 = g补充结算.m_超大额限额公务员补助
        cur医保总费用 = g补充结算.m_医保总费用
        dbl结算总费用 = g补充结算.m_结算总费用
        
        If str疾病编码 = "" Then
            dblHIS总额 = g补充结算.m_HIS总费用
            dbl差额 = dblHIS总额 - dbl结算总费用
        End If
        
        str特殊结算方式 = g补充结算.m_特殊结算方式
        str特殊结算说明 = g补充结算.m_特殊结算说明
        str结算编号 = g补充结算.m_结算编号
        str就诊顺序号 = g补充结算.m_就诊顺序号
    Else
        '调用接口
        If CommServer(IIf(str疾病编码 <> "", "CALSPECCLIN", "CALCLIN")) = False Then Exit Function
    
            
        cur全自付 = Val(GetElemnetValue("FEEOUT"))
        cur挂钩自付 = Val(GetElemnetValue("FEESELF"))
        cur起付线 = Val(GetElemnetValue("STARTFEE"))
        cur基数自付 = Val(GetElemnetValue("ENTERSTARTFEE"))
        
        If str人员类别 = "32" Or str人员类别 = "34" Then
            cur统筹支付 = Val(GetElemnetValue("ALLOWFUND"))
            cur大病统筹 = 0
        Else
            cur统筹支付 = Val(GetElemnetValue("FUND1PAY"))
            cur大病统筹 = Val(GetElemnetValue("FUND2PAY"))
        End If
        cur统筹自付 = Val(GetElemnetValue("FUND1SELF"))
        cur大病自付 = Val(GetElemnetValue("FUND2SELF"))
        cur超限自付 = Val(GetElemnetValue("FEEOVER"))
        
        cur公务员补助起付标准 = Val(GetElemnetValue("STARTFEE2STD"))
        cur公务员补助起付线 = cur起付线
        cur普通门诊公务员补助累计 = Val(GetElemnetValue("ENTERLMT3"))
        cur公务员补助 = Val(GetElemnetValue("FUND3PAY"))
        cur超大额限额公务员补助 = Val(GetElemnetValue("FUND3OVER"))
        cur医保总费用 = Val(GetElemnetValue("FEEALL"))
        dbl结算总费用 = Val(GetElemnetValue("CALFEEALL"))
        
        If str疾病编码 = "" Then
            dblHIS总额 = Val(GetElemnetValue("HOSPFEEALL"))
            dbl差额 = dblHIS总额 - dbl结算总费用
        End If
        
        str特殊结算方式 = GetElemnetValue("SPECCALFLAG")
        str特殊结算说明 = GetElemnetValue("SPECCALFLAGTXT")
        str结算编号 = GetElemnetValue("BALANCEID")
        str就诊顺序号 = GetElemnetValue("BILLNO")
        
    End If
    
    Call SaveBalanceLog(1, lng结帐ID, lng病人ID, GetElemnetValue("BILLNO"), str结算编号, IIf(str疾病编码 <> "", "18", "11"))
    
    If str疾病编码 <> "" Then
        str就诊顺序号 = "特殊" & str就诊顺序号 '把疾病编码与就诊顺序号连在一起
    Else
        str就诊顺序号 = "普通" & str就诊顺序号         '表示普通门诊
    End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_贵阳市, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
                
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_贵阳市 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 + cur统筹支付 + cur统筹自付 + cur基数自付 + cur超限自付 + cur大病统筹 + cur大病自付 & "," & _
        cur统筹报销累计 + cur统筹支付 + cur大病统筹 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_贵阳市 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户余额 & "," & cur帐户余额 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & ",0," & cur基数自付 & "," & cur发生费用 & "," & _
        cur全自付 & "," & cur挂钩自付 & "," & cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & "," & cur大病自付 & "," & cur超限自付 & "," & _
        cur个人帐户 & ",'" & str结算编号 & "',null,null,'" & str就诊顺序号 & "',0,'" & AnalyseComputer & "','" & gstrVersion & "','" & IIf(str疾病编码 <> "", "18", "11") & "','" & Mid(str就诊顺序号, 3) & "'," & _
            "NULL,'" & str疾病编码 & "','" & str保险类别 & "',to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    '门诊不存在设置清算方式，所以清算方式置零（正确的取值范围是从1开始）
      '20110812周玉强增加工信息，将gstr工伤认定编号改为gstr工伤信息
    gstrSQL = "zl_结算附加信息_Insert (" & lng结帐ID & "," & cur公务员补助起付标准 & "," & cur公务员补助起付线 & "," & cur普通门诊公务员补助累计 & "," & cur公务员补助 & "," & cur超大额限额公务员补助 & ",0,0," & _
        "'" & mstr单病种编码 & "'," & mint结算方式 & ",NULL,0," & dbl结算总费用 & "," & cur医保总费用 & ",'" & gstr工伤信息 & "',0,'" & gstr处方本号 & "','" & str特殊结算方式 & "','" & str特殊结算说明 & "')"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    '---------------------------------------------------------------------------------------------
    
    '后台强行进行医保校正，程池富 2011-04-05 Begin
    If str人员类别 = "32" Or str人员类别 = "34" Then
        
        str结算方式 = "医保基金|" & Format(cur统筹支付, "0.00")
    Else
        str结算方式 = "个人帐户|" & Format(cur个人帐户, "0.00") & _
                    "||医保基金|" & Format(cur统筹支付, "0.00") & _
                    "||大病统筹|" & Format(cur大病统筹, "0.00")
    End If
    str结算方式 = str结算方式 & "||医疗补助|" & Format(cur公务员补助, "0.00")
    If mint结算方式 = 1 And dbl差额 > 0 Then
        str结算方式 = str结算方式 & "||差额记帐|" & Format(dbl差额, "0.00")
    End If
    gstrSQL = "select sum(实收金额) from " & IIf(gblnHIS1026 = True, "门诊费用记录", "病人费用记录") & " where 结帐ID=[1]"
    dblHIS总额 = zlDatabase.OpenSQLRecord(gstrSQL, "结账HIS总费用", lng结帐ID).Fields(0)
    dbl现金 = dblHIS总额 - cur统筹支付 - cur大病统筹 - cur公务员补助 - cur个人帐户
    If gblnHIS1029 Then
        If Not bln挂号 Then
            strAdvance = str结算方式
        End If
    Else
        If Not bln挂号 Then
            '门诊部份由接口内部直接校正，不再由费用系统校正
            '校对不需返回现金
            'str结算方式 = str结算方式 & "||现金|" & dbl现金
            'gstrSQL = "zl_门诊收费结算_Update(" & lng结帐ID & ",'" & "" & "'," & 0 & ",'" & str结算方式 & "'," & 0 & ")"
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
            gcnOracle.Execute gstrSQL
        Else
            '挂号部份由挂号程序校正，接口不用处理
        End If
        
        '结算金额累加
        g门诊数据.dbl大病基金 = g门诊数据.dbl大病基金 + cur大病统筹
        g门诊数据.dbl费用总额 = g门诊数据.dbl费用总额 + dblHIS总额
        g门诊数据.dbl个人帐户 = g门诊数据.dbl个人帐户 + cur个人帐户
        g门诊数据.dbl公务员补助 = g门诊数据.dbl公务员补助 + cur公务员补助
        g门诊数据.dbl现金 = g门诊数据.dbl现金 + dbl现金
        g门诊数据.dbl医保基金 = g门诊数据.dbl医保基金 + cur统筹支付
        
        If mint结算方式 = 1 And dbl差额 > 0 Then '单病种差额才加以报销。
            g门诊数据.dbl差额记帐 = g门诊数据.dbl差额记帐 + dbl差额
        Else '如果存在差额但不是单病种，则将差额累计到现金部分。
            g门诊数据.dbl现金 = g门诊数据.dbl现金 + dbl差额
        End If
        
        '最后一张单据显示最终的找补情况' And intMax > 1
        If intMAX = intCurr And bln挂号 = False Then
            Call frm门诊找零贵阳.ShowForm(intMAX)
        End If
        strAdvance = "" '接口内部已经校正，不再需要HIS校正
    End If
    
    门诊结算_贵阳 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function


Public Function 门诊结算冲销_贵阳(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    'Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Dim str结算编号 As String, str就诊顺序号 As String, curDate As Date, rsTemp As New ADODB.Recordset
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency, int住院次数累计 As Integer
    Dim lng冲销ID As Long
    Dim bln离休 As Boolean
    Dim str支付类型 As String
    
    On Error GoTo errHandle
    
    '退费
    '判断是否有结帐记录，如果有说明是住院结帐实现的
    gstrSQL = "Select 1 from 病人结帐记录 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否有结帐记录，如果有说明是住院结帐实现的", lng结帐ID)
    If rsTemp.RecordCount = 0 Then
        gstrSQL = "select distinct A.结帐ID from " & IIf(gblnHIS1026 = True, "门诊费用记录", "病人费用记录") & " A," & _
             IIf(gblnHIS1026 = True, "门诊费用记录", "病人费用记录") & " B " & _
                  " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Else
        gstrSQL = "select distinct A.ID AS 结帐ID from 病人结帐记录 A,病人结帐记录 B " & _
            " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊退费", lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊退费", TYPE_贵阳市, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    lng病人ID = rsTemp!病人ID
    cur个人帐户 = Nvl(rsTemp!个人帐户支付, 0)
    str结算编号 = Nvl(rsTemp("支付顺序号"), "")
    str就诊顺序号 = Nvl(rsTemp("备注"), "")
    If str就诊顺序号 = "" Then
        Err.Raise 9000, gstrSysName, "该单据没有保存就诊顺序号，不能做废。"
        Exit Function
    End If
'    If Left(str就诊顺序号, 2) = "特殊" Then
'        MsgBox "目前不支持特殊门诊的作废。", vbInformation, gstrSysName
'        Exit Function
'    End If
    str支付类型 = Nvl(rsTemp!医疗类别)
    If str支付类型 = "" Then str支付类型 = IIf(Left(str就诊顺序号, 2) = "特殊", "18", "11")
    str就诊顺序号 = Mid(str就诊顺序号, 3)
    curDate = zlDatabase.Currentdate
    
    'If Get验证_贵阳(0, str卡号, str医保号, str分中心编号, str密码, lng病人ID, True) = False Then Exit Function
    Call 多单据收费_退费(lng结帐ID)
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", str就诊顺序号)     ' 就诊顺序号
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str结算编号)    ' 结算编号
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str支付类型)   ' 支付类别
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)    ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(curDate, "yyyy-MM-dd HH:mm:ss"))  ' 办理日期
    
    '调用接口
    bln离休 = IS离休(lng病人ID)
    If MsgBox("是否退票？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    If CommServer("RETBALANCE", IIf(bln离休, 1, 0)) = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(TYPE_贵阳市, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_贵阳市 & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_贵阳市 & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & rsTemp("发生费用金额") * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") & "," & _
        cur个人帐户 * -1 & ",'" & str结算编号 & "',null,null,'" & Nvl(rsTemp("备注")) & "'," & _
        "0,'" & AnalyseComputer & "','" & gstrVersion & "','" & str支付类型 & "','" & Nvl(rsTemp!就诊流水号) & "'," & _
        "NULL,'" & Nvl(rsTemp!病种名称) & "','" & Nvl(rsTemp!并发症) & "',to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    gstrSQL = "Select * From 结算附加信息 Where 结帐ID=" & lng结帐ID
    Call OpenRecordset_OtherBase(rsTemp, "提取结算附加记录", gstrSQL, gcnGYYB)
    If rsTemp.RecordCount <> 0 Then
        gstrSQL = "zl_结算附加信息_Insert (" & lng冲销ID & "," & -1 * Nvl(rsTemp!公务员补助起付标准, 0) & "," & -1 * Nvl(rsTemp!公务员补助起付线, 0) & "," & -1 * Nvl(rsTemp!普通门诊公务员补助累计, 0) & "," _
            & -1 * Nvl(rsTemp!公务员补助, 0) & "," & -1 * Nvl(rsTemp!超大额公务员补助, 0) & ",0,0,'" & Nvl(rsTemp!单病种编码_结算) & "'," & Nvl(rsTemp!结算方式, 0) & ",'" & Nvl(rsTemp!单病种) & "'," & _
            Nvl(rsTemp!清算方式, 0) & "," & -1 * Nvl(rsTemp!结算总费用, 0) & "," & -1 * Nvl(rsTemp!医保总费用, 0) & ",'" & Nvl(rsTemp!工伤认定编号) & "',0,'" & Nvl(rsTemp!处方本号) & "')"
        gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    门诊结算冲销_贵阳 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 个人帐户转预交_贵阳(lng预交ID As Long, cur个人帐户 As Currency, strSelfNo As String, str顺序号 As String, ByVal lng病人ID As Long) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    
    个人帐户转预交_贵阳 = False
End Function

Public Function 入院登记_贵阳(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim str卡号 As String, str分中心编号 As String, str密码 As String, str人员类别 As String
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str疾病编码 As String
    Dim strTemp As String, str提示 As String, str诊断 As String, lng参保前在院 As Long
    Dim str支付类型 As String, dbl帐户余额 As Double
    Dim str特殊结算方式 As String, str清算方式 As String, str清算病种编码 As String, str清算病种名称 As String, str保险类别 As String, str特殊结算说明 As String
    On Error GoTo errHandle
    
    'If Get验证_贵阳(1, str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
    
    'Modified by ZYB 20080703 接口文档取消本参数
'    '判断该病人是否参保前在院
'    lng参保前在院 = 0
'    If Get保险参数_贵阳("入院时选择参保前在院") = "1" Then
'        If MsgBox("该病人参保前是否已经在院？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
'            lng参保前在院 = 1
'        End If
'    End If
    
    '获得病人出院诊断
    gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=[1] and A.主页ID=[2]" & _
              " and A.诊断类型=1 and A.诊断次序=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "入院登记", lng病人ID, lng主页ID)
    If rsTemp.EOF = False Then
        str诊断 = ToVarchar(IIf(IsNull(rsTemp("描述信息")), "疾病", rsTemp("描述信息")), 128)
    Else
        gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=[1] and A.主页ID=[2]" & _
                  " and A.诊断类型=2 and A.诊断次序=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "入院登记", lng病人ID, lng主页ID)
        If rsTemp.RecordCount <> 0 Then
            str诊断 = ToVarchar(Nvl(rsTemp!描述信息, "疾病"), 128)
        Else
            str诊断 = "疾病"   '诊断不论如何不能为空
        End If
    End If
    
    '获得其它入院信息
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.入院方式,nvl(A.二级院转入,0) as 二级院转入,A.门诊医师,A.入院日期,A.入院病床,B.名称 as 入院科室," & _
              "     C.住院号,D.保险类别,D.生育标志,D.帐户余额,D.支付类型 " & _
              " from 病案主页 A,部门表 B,病人信息 C,保险帐户 D " & _
              " Where A.病人ID=D.病人ID And D.险类=[1] And A.病人ID=C.病人ID and A.入院科室ID = B.ID And A.病人ID =[2] And A.主页ID = [3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "入院登记", TYPE_贵阳市, lng病人ID, lng主页ID)
    dbl帐户余额 = Nvl(rsTemp!帐户余额, 0)
    str保险类别 = Nvl(rsTemp!保险类别)
    str支付类型 = Nvl(rsTemp!支付类型, "31")
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' 卡类别
    Call InsertChild(mdomInput.documentElement, "CARDDATA", mstr卡号)           ' 磁卡编码
    Call InsertChild(mdomInput.documentElement, "PASSWORD", mstr密码)         ' 密码
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstr医保号)      ' 个人编号
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' 社会保障号
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP地址
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM卡芯片
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", str保险类别)   ' 保险类别
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str支付类型)
    Call InsertChild(mdomInput.documentElement, "HOSPNO", ToVarchar(rsTemp("住院号"), 20))     ' 住院号
'    Call InsertChild(mdomInput.documentElement, "ISINHOSP", lng参保前在院)     ' 参保前已在院 1：是；0：否
    Call InsertChild(mdomInput.documentElement, "DIAGNOSES", str诊断) ' 诊断
    Call InsertChild(mdomInput.documentElement, "DOCTOR", ToVarchar(rsTemp("门诊医师"), 20)) ' 诊断医生
    Call InsertChild(mdomInput.documentElement, "OFFICE", ToVarchar(rsTemp("入院科室"), 20)) ' 科室
    Call InsertChild(mdomInput.documentElement, "REGDATE", Format(rsTemp("入院日期"), "yyyy-MM-dd HH:mm:ss")) ' 入院时间
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.姓名) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))  ' 办理日期
    Call InsertChild(mdomInput.documentElement, "GSRDBH", gstr工伤认定编号)  ' 工伤认定编号
    Call InsertChild(mdomInput.documentElement, "KFZYBZ", gint工伤康复住院)  ' 工伤康复住院
    
    '调用接口
    If CommServer("HOSPREG") = False Then Exit Function
    
    Dim int住院次数累计 As Integer
    Dim cur本次起付线 As Currency
    Dim cur起付线累计 As Currency
    Dim cur基本统筹限额 As Currency
    Dim cur统筹报销累计 As Currency
    Dim cur大额统筹限额 As Currency
    Dim cur大额统筹累计 As Currency
    Dim str封锁信息 As String
    
    int住院次数累计 = Val(GetElemnetValue("HOSPTIMES"))
    cur本次起付线 = Val(GetElemnetValue("STARTFEE"))
    cur起付线累计 = Val(GetElemnetValue("STARTFEEPAID"))
    cur基本统筹限额 = Val(GetElemnetValue("FUND1LMT"))
    cur统筹报销累计 = Val(GetElemnetValue("FUND1PAID"))
    cur大额统筹限额 = Val(GetElemnetValue("FUND2LMT"))
    cur大额统筹累计 = Val(GetElemnetValue("FUND2PAID"))
    
    str封锁信息 = GetElemnetValue("LOCKINFO")
    Do Until str封锁信息 = ""
        strTemp = Left(str封锁信息, 2)
        str封锁信息 = Mid(str封锁信息, 41)
        
        str提示 = str提示 & Switch(strTemp = "11", "、待遇审核期", strTemp = "21", "、卡封锁", strTemp = "31", "、基本统筹欠费", _
                                   strTemp = "32", "、大额统筹未缴费", strTemp = "41", "、停保", strTemp = "51", "、退保")
    Loop
    str提示 = str提示 & GetElemnetValue("NOTE")
    If str提示 <> "" Then
        MsgBox "请注意该医保病人情况：" & Mid(str提示, 2) & "。", vbInformation, gstrSysName
    End If
    '<SPECCALFLAG>特殊结算标志</SPECCALFLAG>
    '<RECKONINGTYPE>清算方式</RECKONINGTYPE>
    '<SINGLEILLNESSCODE>单病种编码</SINGLEILLNESSCODE>
    str特殊结算方式 = GetElemnetValue("SPECCALFLAG")
    str特殊结算说明 = GetElemnetValue("SPECCALFLAGTXT")
    str清算方式 = GetElemnetValue("RECKONINGTYPE")
    str清算病种编码 = GetElemnetValue("SINGLEILLNESSCODE")
    str清算病种名称 = GetElemnetValue("SINGLEILLNESSNAME")
    If str特殊结算方式 = "" Then str特殊结算方式 = "00"
    
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_贵阳市 & "," & Year(datCurr) & "," & _
        dbl帐户余额 & ",0,0," & cur统筹报销累计 & "," & int住院次数累计 & "," & cur本次起付线 & "," & cur起付线累计 & _
         "," & cur基本统筹限额 & "," & cur大额统筹限额 & "," & cur大额统筹累计 & ",'" & ToVarchar(str提示, 100) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_贵阳市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'顺序号','''" & GetElemnetValue("BILLNO") & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存入院登记顺序号")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'工伤认定编号','''" & gstr工伤认定编号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存工伤认定编号")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'工伤康复住院','" & gint工伤康复住院 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存工伤康复住院标志")
    '保存设置的清算方式
    gstrSQL = "ZL_医保病人住院信息_INSERT(" & _
              lng病人ID & "," & lng主页ID & ",'" & gstrUserName & "',2," & str保险类别 & ",'" & str清算病种名称 & "',NULL,NULL," & _
              "NULL,NULL,NULL,NULL,NULL,'" & str清算病种编码 & "',NULL,NULL,NULL,NULL,'" & str清算方式 & "','" & str特殊结算方式 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存入院函数返回的清算方式")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'单病种','''" & str清算病种编码 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存单病种编码")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'清算方式','''" & str清算方式 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存单病种编码")
    
    入院登记_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_贵阳(lng病人ID As Long, lng主页ID As Long, Optional ByVal bln出院 As Boolean = False) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
            '取入院登记验证所返回的顺序号
            
    '修改说明
    '时间：2005-01-14
    '修改人：朱玉宝
    '修改内容：出院登记接口增加入参—ICD编码，也单独提供了上传ICD编码的接口，与齐凯联系后，暂定本接口不上传ICD编码，由上传ICD编码完成
    
    Dim str医保号 As String
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str诊断 As String, str其它诊断 As String
    Dim str病案号 As String, str出院转归 As String, lngPos As Long
    Dim str出院日期 As String
    
    On Error GoTo errHandle
    
    If mbln医保出院 Or bln出院 Then
        '由于医保要求出院日期必须大于结算日期，而多数医院是先出院后结算，因此，在调医保出院函数时，取最后一次结算的日期+1秒做为出院日期传过去
        gstrSQL = "SELECT to_Char(收费时间,'yyyy-MM-dd hh24:mi:ss') AS 收费时间 FROM 病人结帐记录 " & _
                 " WHERE ID=( " & _
                 "     SELECT MAX(记录ID)  " & _
                 "     FROM 保险结算记录  " & _
                 "     WHERE 病人ID=[1] AND 险类=[2])"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取最后一次结帐的时间", lng病人ID, TYPE_贵阳市)
        If rsTemp.RecordCount <> 0 Then str出院日期 = Format(DateAdd("s", 1, rsTemp!收费时间), "yyyy-MM-dd HH:mm:ss")
        
        '从数据库中读出已存储的值
        gstrSQL = "select 卡号,医保号,顺序号 from 保险帐户 where 病人ID=[1] and 险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院登记", lng病人ID, TYPE_贵阳市)
        
        str医保号 = IIf(IsNull(rsTemp("医保号")), "", rsTemp("医保号"))
        
        '获得病人出院信息
        gstrSQL = "SELECT A.出院方式,nvl(C.病案号,B.住院号) AS 病案号  " & _
                 " FROM 病案主页 A,病人信息 B,住院病案记录 C " & _
                 " WHERE A.病人ID=[1] AND A.主页id=[2] AND A.病人id=B.病人id AND A.病人id=C.病人id(+)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院登记", lng病人ID, lng主页ID)
        str病案号 = Nvl(rsTemp("病案号"), lng病人ID)
        Select Case rsTemp("出院方式")
            Case "正常", "治愈"
                str出院转归 = "1"
            Case "好转"
                str出院转归 = "2"
            Case "死亡"
                str出院转归 = "3"
            Case Else
                str出院转归 = "9"
        End Select
        
        '获得病人出院诊断
        gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=[1] and A.主页ID=[2]" & _
                  " and A.诊断类型=3 and A.诊断次序=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院登记", lng病人ID, lng主页ID)
        If rsTemp.EOF = False Then
            str诊断 = Nvl(rsTemp("描述信息"), "疾病")
            '将不同形式的分隔符统一
            str诊断 = Replace(str诊断, "，", ",")
            str诊断 = Replace(str诊断, "；", ",")
            str诊断 = Replace(str诊断, "、", ",")
            str诊断 = Replace(str诊断, ";", ",")
            lngPos = InStr(str诊断, ",")
            If lngPos > 0 Then
                str其它诊断 = Mid(str诊断, lngPos + 1)
                str诊断 = Mid(str诊断, 1, lngPos - 1)
            End If
        Else
            str诊断 = "疾病"   '诊断不论如何不能为空
        End If
            
        '获得其它出院信息
        datCurr = zlDatabase.Currentdate
        gstrSQL = "select A.住院医师,A.出院日期,A.出院病床,B.名称 as 出院科室 from 病案主页 A,部门表 B " & _
                 " Where A.出院科室ID = B.ID And A.病人ID =[1] And A.主页ID = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院登记", lng病人ID, lng主页ID)
        If str出院日期 = "" Then str出院日期 = Format(rsTemp!出院日期, "yyyy-MM-dd HH:mm:ss")
        If str出院日期 = "" Then str出院日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        '对XML DomDocument对象进行初始化
        If InitXML = False Then Exit Function
        Call InsertChild(mdomInput.documentElement, "PERSONCODE", str医保号)     ' 个人编码
        Call InsertChild(mdomInput.documentElement, "DOCNO", str病案号)          ' 病案号
        Call InsertChild(mdomInput.documentElement, "DIAGNOSES", ToVarchar(str诊断, 128))          ' 诊断
        Call InsertChild(mdomInput.documentElement, "OTHERDIAGNOSES", ToVarchar(str其它诊断, 128)) ' 其他诊断
        Call InsertChild(mdomInput.documentElement, "OUTTYPE", str出院转归)                        ' 转归类别
        Call InsertChild(mdomInput.documentElement, "ICD", "")                       ' ICD疾病编码
        Call InsertChild(mdomInput.documentElement, "DOCTOR", ToVarchar(Nvl(rsTemp("住院医师"), "ZLHIS"), 20))  ' 诊断医生
        Call InsertChild(mdomInput.documentElement, "OFFICE", ToVarchar(rsTemp("出院科室"), 20))   ' 科室
        Call InsertChild(mdomInput.documentElement, "REGDATE", str出院日期) ' 出院日期
        Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.姓名) ' 操作员
        Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))  ' 办理日期
        
        '调用接口
        If CommServer("HOSPOUT") = False Then Exit Function
    End If
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_贵阳市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    MsgBox IIf(mbln医保出院 Or bln出院, "成功办理HIS、医保出院！", "仅办理了HIS出院！"), vbInformation, gstrSysName
    出院登记_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_贵阳(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim str顺序号 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    gstrSQL = " Select 顺序号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取就诊流水号", TYPE_贵阳市, lng病人ID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有找到该病人的医保档案！", vbInformation, gstrSysName
        Exit Function
    End If
    str顺序号 = Nvl(rsTemp!顺序号)

    '此处不对医保调用成功与否进行检查，因出院时可能仅办理了HIS出院
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", str顺序号) ' 入院时间
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.姓名) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss"))  ' 办理日期
    Call CommServer("RETHOSPOUT")

    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_贵阳市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")

    出院登记撤销_贵阳 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_贵阳(rsExse As Recordset, ByVal lng病人ID As Long, Optional ByVal bln结帐 As Boolean = True) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim cn上传 As New ADODB.Connection, rsTemp As New ADODB.Recordset, rs病种 As ADODB.Recordset, rs诊断 As New ADODB.Recordset, rs对码 As New ADODB.Recordset
    Dim lng病种ID As Long, lng主页ID As Long, str项目编码 As String
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, str人员类别 As String
    Dim cur个人帐户 As Double, cur统筹支付 As Double, cur大病统筹 As Double, cur发生费用 As Double
    Dim dbl结算总费用  As Double, dblHIS总费用 As Double, dbl差额 As Double
    Dim cur公务员补助 As Double, cur医疗照顾公务员补助 As Double, cur超大额公务员补助 As Double, int生育标志 As Integer
    Dim str医生 As String, str科室 As String, str特殊药品 As String
    Dim bln特殊药品 As Boolean, bln特殊药品审批状态 As Boolean, bln存在未审批的特殊药品 As Boolean
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim rsTmp       As ADODB.Recordset
    On Error GoTo errHandle
    mlng病人ID = 0         '初始化。只要一选择病人，就会调用本过程，也就会设成0
    
    If rsExse.RecordCount = 0 Then
        ShowMsgbox "该病人没有有发生费用，无法进行结算操作。"
        Exit Function
    End If
    
    rsExse.MoveFirst
    '打开另外一个连接串，以达到不受当前连接事务的控制
    Set cn上传 = GetNewConnection
    '此处密码确定是得不到的，所以要强制刷卡
    Screen.MousePointer = vbDefault
    
    '取该病人的基本信息
    gstrSQL = " Select A.人员身份,A.工伤认定编号,A.工伤康复住院,B.编码,C.住院次数 from 保险帐户 A,保险病种 B,病人信息 C" & _
              " Where A.病人ID=[1] And A.病人ID=C.病人ID and A.险类=[2]  and A.病种ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "住院预算", lng病人ID, TYPE_贵阳市)
    If rsTemp.EOF = False Then
        lng主页ID = rsTemp!住院次数
        gstr工伤认定编号 = Nvl(rsTemp!工伤认定编号)
        gint工伤康复住院 = Nvl(rsTemp!工伤康复住院, 0)
        str人员类别 = Nvl(rsTemp("人员身份"), "")
        '转换人员身份
        str人员类别 = Switch(str人员类别 = "在职", "11", str人员类别 = "退休", "21", _
                          str人员类别 = "省属离休", "32", str人员类别 = "市属离休", "34", _
                          str人员类别 = "普通居民", "41", str人员类别 = "低保对象", "42", _
                          str人员类别 = "三无人员", "43", str人员类别 = "低收家庭", "44", _
                          str人员类别 = "重度残疾", "45", True, "11") '界面上仍显示低收入家庭，但数据库只有8位，只能保存为低收家庭
    End If
    
    mstr卡号 = ""
    mstr密码 = ""
    If Get验证_贵阳(1, str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
    Screen.MousePointer = vbHourglass
    
    mbln无密码结算 = False
    If bln结帐 Then
        If MsgBox("是否进行无密码结算（仅适用于病人逃费的情况）？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mbln无密码结算 = True
    End If
    
    '病历录入检查
    If gbln批量虚拟结算 = False And str密码 = "1" Then
        If ging病历检查 <> 0 Then
            '0不检查,1检查提醒,2禁止结帐
            gstrSQL = "Select 备注 From 病历登记记录_贵阳 Where 险类=[1] And 病人ID=[2] And 主页ID=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSQL, TYPE_贵阳市, lng病人ID, lng主页ID)
            If rsTemp.RecordCount = 0 Then
                If ging病历检查 = 1 Then
                    If MsgBox("该病人未登记病历记录，是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
                Else
                    ShowMsgbox "该病人未登记病历记录，请先登记！"
                    Exit Function
                End If
            Else
                '把病历登记的备注情况显示出来
                ShowMsgbox "该病人有以下备注信息，请核实：" & vbNewLine & Nvl(rsTemp!备注)
            End If
        End If
    End If
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' 卡类别
    Call InsertChild(mdomInput.documentElement, "CARDDATA", str卡号)                 ' 磁条数据
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                     ' 社会保障号
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP地址
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM卡芯片
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", str医保号)     ' 个人编码
    Call InsertChild(mdomInput.documentElement, "ISCAL", 0)         ' 是否结算
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", "0")     ' 账户支付额
    Call InsertChild(mdomInput.documentElement, "INVOICENO", "") ' 发票号
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' 消费明细
    
    rsExse.Sort = " 登记时间 asc"
    Do Until rsExse.EOF
        '检查是否特殊项目
        bln特殊药品 = False                     '当前是否特殊药品
        bln特殊药品审批状态 = False             '当前特殊药品是否审批
        
        If IIf(IsNull(rsExse("是否上传")), "0", rsExse("是否上传")) = "0" And Nvl(rsExse!金额, 0) <> 0 Then
            gstrSQL = "SELECT C.药品ID,C.规格,E.名称 AS 剂型  FROM 药品目录 C,药品信息 D,药品剂型 E WHERE C.药品ID=" & rsExse("收费细目ID") & " AND C.药名ID=D.药名ID AND D.剂型=E.编码"
            gstrSQL = "select A.类别,A.名称,B.项目编码,nvl(A.规格,F.规格) AS 规格,F.剂型,A.计算单位 from 收费细目 A,保险支付项目 B,(" & gstrSQL & _
                    ") F where A.ID=[1] and A.ID=B.收费细目ID  AND A.Id=F.药品ID(+) and B.险类=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "住院预算", CLng(rsExse("收费细目ID")), TYPE_贵阳市)
            If rsTemp.EOF = True Then
                ShowMsgbox "有项目未设置医保编码，不能结算。"
                Exit Function
            End If
            
            '只上传只传递过的数据
            str医生 = ToVarchar(IIf(IsNull(rsExse("医生")), UserInfo.姓名, rsExse("医生")), 20)
            str科室 = ToVarchar(IIf(IsNull(rsExse("开单部门")), UserInfo.部门, rsExse("开单部门")), 56)
            
             On Error Resume Next
            '20100301 BY ZYB
            '针对特殊药品项目，如果未审批则不上传，也不汇总HIS总金额，以便程序后面的核对总金额能顺利通过
            '如果有特殊药品项目未审批的，给予提示，允许继续向下执行
            str项目编码 = ""
            str项目编码 = Nvl(rsExse!保险编码)      '缺省以当前对码数据为准
            
            If mbln特殊药品提示 Then
                gstrSQL = " Select b.标志 From " & IIf(gblnHIS1026 = True, "住院费用记录", "病人费用记录") & " a,特殊药品收费 b"
                gstrSQL = gstrSQL & vbCrLf & "Where a.ID = b.费用id And a.病人ID = [1] And a.主页ID = [2]"
                Set rs对码 = zlDatabase.OpenSQLRecord(gstrSQL, "读取特殊药品", lng病人ID, lng主页ID)
                
                gstrSQL = " Select 1 From 收费细目 Where 说明 Is Not NULL And 类别 IN ('5','6','7') And ID=[1]"
                Set rs对码 = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目信息", CLng(rsExse!收费细目ID))
                If rs对码.RecordCount <> 0 Then
                    bln特殊药品 = True
                    gstrSQL = " Select 标志 From 特殊药品收费 " & _
                              " Where 病人ID=[1] And 主页ID=[2] And 费用ID=" & _
                              " (Select ID from " & IIf(gblnHIS1026 = True, "住院费用记录", "病人费用记录") & " where 记录性质=[1] and 记录状态=[2] and No=[3] and 序号=[4])"
                              
                    Set rs对码 = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目信息", lng病人ID, lng主页ID, CLng(rsExse!记录性质), CLng(rsExse!记录状态), CStr(rsExse!NO), CLng(rsExse!序号))
                    If rs对码.RecordCount <> 0 Then
                        bln特殊药品审批状态 = True
                        If Nvl(rs对码!标志, 0) = 0 Then
                           '西药全自费编码：81900090009
                            '20110201开始自费西药编码:   810851900099,周玉强修改
                            '20110201开始自费中成药编码: 820851900099,周玉强修改
                            '中成药及中草药全自费编码:   829000900099,周玉强修改
                            If rsExse!收费类别 = "5" Then
                            str项目编码 = "810851900099"
                            ElseIf rsExse!收费类别 = "6" Then
                            str项目编码 = "820851900099"
                            Else
                            str项目编码 = "829000900099"
                            End If
                           ' str项目编码 = IIf(rsExse!收费类别 = "5", "81900090009", "82900090009")
                        End If
                    Else
                        bln存在未审批的特殊药品 = True  '是否存在未审批的特殊药品
                    End If
                End If
            End If
            
            Err = 0
            On Error GoTo errHandle
            
            '如果是婴儿费，则传自费编码
            If Nvl(rsExse!婴儿费, 0) <> 0 Then
               '20110201开始自费西药编码:   810851900099,周玉强修改
                 '20110201开始自费中成药编码: 820851900099,周玉强修改
                 '中成药及中草药全自费编码:   829000900099 ,周玉强修改
                '诊疗全自费编码：34900099
                If rsExse!收费类别 = "5" Then
                    str项目编码 = "810851900099"
                ElseIf rsExse!收费类别 = "6" Then
                     str项目编码 = "820851900099"
                ElseIf rsExse!收费类别 = "7" Then
                    str项目编码 = "829000900099"
                Else
                    str项目编码 = "34900099"
                End If
            Else
                '如果保险编码为空，则需要用户选择编码
                If str项目编码 = "" Then
                    str项目编码 = GetItemInsure_贵阳(lng病人ID, rsExse!收费细目ID, False)
                End If
                If str项目编码 = "" Then str项目编码 = Nvl(rsExse!医保项目编码)
            End If
            
            If Not bln特殊药品 Or (bln特殊药品 And bln特殊药品审批状态) Then
                Set nodRow = InsertChild(nodRowset, "ROW", "")
                Call nodRow.setAttribute("ITEMSERIAL", ToVarchar(rsExse("NO") & "_" & rsExse("序号") & "_" & rsExse("记录性质") & "_" & rsExse("记录状态"), 20)) '数据批号，用于唯一代表数据
                Call nodRow.setAttribute("ITEMCODE", ToVarchar(str项目编码, 12))
                Call nodRow.setAttribute("ITEMNAME", ToVarchar(rsExse("收费名称"), 72))
                Call nodRow.setAttribute("SUBJECT", Subject(rsTemp("类别")))
                Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rsTemp("规格"), 40))
                Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rsTemp("剂型"), 20))
                Call nodRow.setAttribute("UNIT", ToVarchar(rsTemp("计算单位"), 20))
                Call nodRow.setAttribute("PRICE", Format(rsExse("价格"), "0.0000"))
                Call nodRow.setAttribute("QUANTITY", Format(rsExse("数量"), "0.00"))
                Call nodRow.setAttribute("FROMOFFICE", str科室)   '开单科室
                Call nodRow.setAttribute("FROMDOCT", str医生)     '开单医生
                Call nodRow.setAttribute("TOOFFICE", str科室)    '受单科室
                Call nodRow.setAttribute("TODOCT", str医生)      '受单医生
                '处理时间时，为了保证同一保险项目的的收费时间不同，因此在登记时间上按序号加上秒数
                Call nodRow.setAttribute("DODATE", Format(DateAdd("s", rsExse("序号") + IIf(rsExse!记录状态 = 2, 1, 0), rsExse("登记时间")), "yyyy-MM-dd HH:mm:ss")) '办理日期
                Call nodRow.setAttribute("NOTE", ToVarchar(rsExse("摘要"), 512))     '备注
            End If
        End If
        
        If Not bln特殊药品 Or (bln特殊药品 And bln特殊药品审批状态) Then
            cur发生费用 = cur发生费用 + Round(rsExse("金额"), 2)
        Else
            'J12345671100001
            str特殊药品 = str特殊药品 & "," & rsExse("NO") & rsExse("记录性质") & rsExse("记录状态") & String(5 - Len(CStr(rsExse("序号"))), "0") & rsExse("序号")
        End If
        'XieRong 2010-10-22 更新医保编码
        If Nvl(rsExse!医保项目编码) = "" Then
            Dim rsID        As ADODB.Recordset
            gstrSQL = "select ID from " & IIf(gblnHIS1026 = True, "住院费用记录", "病人费用记录") & " where NO=[1] and 序号=[2] and 记录性质=[3] and 记录状态=[4]"
            Set rsID = zlDatabase.OpenSQLRecord(gstrSQL, "费用ID", rsExse("NO"), rsExse("序号"), rsExse("记录性质"), rsExse("记录状态"))
            If rsID.RecordCount > 0 Then
                gstrSQL = "ZL_病人费用记录_更新医保(" & rsID!ID & ",0," & TYPE_贵阳市 & ",,'" & str项目编码 & "')"
                zlDatabase.ExecuteProcedure gstrSQL, "更新医保信息"
            End If
        End If
        rsExse.MoveNext
    Loop
    
    '调用接口
    If CommServer("CALHOSP", IIf(mbln无密码结算, "1", "0")) = False Then Exit Function
    '首先强调不能少传，所以等医保服务器正确返回后再打上标记
    str特殊药品 = str特殊药品 & ","
    If rsExse.RecordCount > 0 Then rsExse.MoveFirst
    Do Until rsExse.EOF
        If IIf(IsNull(rsExse("是否上传")), "0", rsExse("是否上传")) = "0" Then
            '为该条费用记录打上上传标志。上传一条处理一条
            If InStr(1, str特殊药品, "," & rsExse("NO") & rsExse("记录性质") & rsExse("记录状态") & String(5 - Len(CStr(rsExse("序号"))), "0") & rsExse("序号") & ",") = 0 Then
                gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
                cn上传.Execute gstrSQL, , adCmdStoredProc
            End If
        End If
        rsExse.MoveNext
    Loop
    
    cur个人帐户 = Val(GetElemnetValue("ACCTPAY"))
    If str人员类别 = "32" Or str人员类别 = "34" Then
        cur统筹支付 = Val(GetElemnetValue("ALLOWFUND"))
    Else
        cur统筹支付 = Val(GetElemnetValue("FUND1PAY"))
    End If
    cur大病统筹 = Val(GetElemnetValue("FUND2PAY"))
    
'    <FUND3PAY>公务员补助支付</FUND3PAY>
'    <CAREPAY>医疗照顾人员特项公务员补助</CAREPAY>
'    <FUND3OVER>超大额限额公务员补助</ FUND3OVER >
'    <BEARINGFLAG>生育标志</BEARINGFLAG>
    cur公务员补助 = Val(GetElemnetValue("FUND3PAY"))
    dbl结算总费用 = Val(Format(GetElemnetValue("CALFEEALL"), "#0.00;-#0.00;0;"))
    dblHIS总费用 = Val(Format(GetElemnetValue("HOSPFEEALL"), "#0.00;-#0.00;0;"))
    dbl差额 = dblHIS总费用 - dbl结算总费用
    
    '先比较总额是否一致
    If gbln批量虚拟结算 = False Then
        If Format(cur发生费用, "#0.00") <> Format(dblHIS总费用, "#0.00") Then
            If Abs(Val(Format(cur发生费用, "#0.00")) - Val(Format(dblHIS总费用, "#0.00"))) > 0.5 Then
                MsgBox "HIS总额与医保接收到的总费用不一致，不允许结算！" & vbCrLf & _
                    "HIS:" & Format(cur发生费用, "#0.00") & Space(10) & "医保:" & Format(dblHIS总费用, "#0.00"), vbInformation, gstrSysName
                Exit Function
            Else
                If MsgBox("HIS总额与医保接收到的总费用不一致，估计是单价精度不足引起的误差，是否结算？" & vbCrLf & _
                    "HIS:" & Format(cur发生费用, "#0.00") & Space(10) & "医保:" & Format(dblHIS总费用, "#0.00"), vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
            End If
        End If
    End If
    
    If gbln批量虚拟结算 = False Then
        If bln存在未审批的特殊药品 Then
            ShowMsgbox "存在未审批的特殊药品，需审批后才能结算！"
        End If
    End If
    
    '保存病人个人帐户余额
    mstr医保号 = str医保号
    mdbl余额 = cur个人帐户
    
    '保存临时数据，为结算操作做准备
    With g结算数据
        .发生费用金额 = cur发生费用
    End With
    Screen.MousePointer = 0
    If gbln批量虚拟结算 = False Then frm贵阳结算信息.Show 1
    
    Dim str保险类别 As String, strMsg As String
    '显示该病人的险种与清算方式，供操作员核对
    '检查是否选择清算方式或结算方式，没有选择缺省为控制线
    gstrSQL = " Select A.操作员, A.日期, NVL(A.有效标志,2) AS 有效标志,Nvl(A.保险类别,B.保险类别) AS 保险类别, A.病种名称, A.开始日期, A.结束日期, " & _
              "        A.单病种编码_结算, A.结算限额, A.统筹包干,A.个人负担, NVL(A.结算方式,'0') AS 结算方式, " & _
              "        A.单病种, A.基本统筹清单标准, A.基本统筹分担比例, A.大额补助清算标准, A.大额补助分担比例, Nvl(A.清算方式,'1') AS 清算方式,NVL(A.特殊结算方式,'00') AS 特殊结算方式 " & _
              " From 医保病人住院信息 A,保险帐户 B,病人信息 C" & _
              " Where B.病人ID=C.病人ID And C.病人ID=A.病人ID(+) And C.住院次数=A.主页ID(+) And C.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否选择清算方式或结算方式", lng病人ID)
    If Mid(rsTemp!特殊结算方式, 2, 1) = "0" Then    '系统固定的话，第二位非零，对于自己设置的才提示出来供操作员检查，否则不检查
        Select Case rsTemp!保险类别
        Case 2
            strMsg = "保险类别：企业离休医疗保险"
        Case 3
            strMsg = "保险类别：机关事业单位基本医疗保险"
        Case 4
            strMsg = "保险类别：企业生育保险"
        Case 5
            strMsg = "保险类别：机关事业单位生育保险"
        Case 6
            strMsg = "保险类别：居民保险"
        Case 7
            strMsg = "保险类别：工伤保险"
        Case Else
            strMsg = "保险类别：企业基本医疗保险"
        End Select
        strMsg = "请在结算前仔细核对该病人的住院信息！" & vbCrLf & strMsg & vbCrLf
        If rsTemp!有效标志 = 1 Then
            '结算方式
            strMsg = strMsg & "结算方式：" & rsTemp!结算方式
            If rsTemp!结算方式 = "1" Then
                strMsg = strMsg & vbCrLf & _
                    "病种信息：（" & Nvl(rsTemp!单病种编码_结算) & "）" & Nvl(rsTemp!病种名称) & vbCrLf & _
                    "结算限额：" & Nvl(rsTemp!结算限额) & "；统筹包干：" & Nvl(rsTemp!统筹包干) & "；个人负担：" & Nvl(rsTemp!个人负担)
            End If
        Else
            '清算方式
            Select Case rsTemp!清算方式
            Case 2
                strMsg = strMsg & "清算方式：重症病种清算"
            Case 3
                strMsg = strMsg & "清算方式：单病种按人次定额包干清算方式"
            Case 4
                strMsg = strMsg & "清算方式：单病种按日定额包干清算方式"
            Case 5
                strMsg = strMsg & "清算方式：生育保险包干清算"
            Case 6
                strMsg = strMsg & "清算方式：单病种包干清算"
            Case Else
                strMsg = strMsg & "清算方式：控制线清算方式（生育保险中为非包干方式）"
            End Select
            If rsTemp!清算方式 <> "1" Then
                strMsg = strMsg & vbCrLf & _
                    "病种信息：（" & Nvl(rsTemp!单病种) & "）" & Nvl(rsTemp!病种名称)
            End If
        End If
    End If
    Dim str诊断 As String, str其它诊断 As String, lngPos As Long
    '获得病人出院诊断
    gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=[1] and A.主页ID=[2]" & _
              " and A.诊断类型=3 and A.诊断次序=1"
    Set rs诊断 = zlDatabase.OpenSQLRecord(gstrSQL, "出院登记", lng病人ID, lng主页ID)
    If rs诊断.EOF = False Then
        str诊断 = Nvl(rs诊断("描述信息"), "疾病")
        '将不同形式的分隔符统一
        str诊断 = Replace(str诊断, "，", ",")
        str诊断 = Replace(str诊断, "；", ",")
        str诊断 = Replace(str诊断, "、", ",")
        str诊断 = Replace(str诊断, ";", ",")
        lngPos = InStr(str诊断, ",")
        If lngPos > 0 Then
            str其它诊断 = Mid(str诊断, lngPos + 1)
            str诊断 = Mid(str诊断, 1, lngPos - 1)
        End If
    Else
        str诊断 = "疾病"   '诊断不论如何不能为空
    End If
    strMsg = strMsg & vbCrLf & "出院诊断：" & str诊断
    If Trim(str其它诊断) <> "" Then strMsg = strMsg & "其它诊断：" & str其它诊断
    '出院方式 Modify By 程池富 2010-01-16 贵阳医学院要求
    Dim rs出院方式 As New ADODB.Recordset
    gstrSQL = "SELECT 出院方式  FROM 病案主页 where 险类=[1] And 病人ID=[2] And 主页ID=[3]"
    Set rs出院方式 = zlDatabase.OpenSQLRecord(gstrSQL, "出院方式", TYPE_贵阳市, lng病人ID, lng主页ID)
    If Nvl(rs出院方式!出院方式, "无") <> "无" Then strMsg = strMsg & vbCrLf & "出院方式：" & Nvl(rs出院方式!出院方式)
    If gbln批量虚拟结算 = False Then ShowMsgbox strMsg
    
    
    '返回预结算结果
    住院虚拟结算_贵阳 = "医保基金;" & cur统筹支付 & ";0"
    If cur个人帐户 <> 0 Then
        住院虚拟结算_贵阳 = 住院虚拟结算_贵阳 & "|个人帐户;" & cur个人帐户 & ";1" '允许修改个人帐户
    End If
    If InStr(1, "4,5", rsTemp!保险类别) <> 0 Then   '生育保险名称改为产前检查费
        住院虚拟结算_贵阳 = 住院虚拟结算_贵阳 & "|产前检查费;" & cur大病统筹 & ";0"
    Else
        住院虚拟结算_贵阳 = 住院虚拟结算_贵阳 & "|大病统筹;" & cur大病统筹 & ";0"
    End If
    住院虚拟结算_贵阳 = 住院虚拟结算_贵阳 & "|医疗补助;" & Format(cur公务员补助, "#0.00;-#0.00;0;") & ";0"
    '只有结算方式为单病种包干的才存在差额记帐，否则归入现金
    If rsTemp!有效标志 = 1 And rsTemp!结算方式 = "单病种包干结算" Then 住院虚拟结算_贵阳 = 住院虚拟结算_贵阳 & "|差额记帐;" & dbl差额 & ";0"
    mlng病人ID = lng病人ID  '表示该病人已经进行了虚拟结算
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_贵阳(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp   As ADODB.Recordset
    Dim lng主页ID As Long
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double, cur帐户余额 As Currency
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, cur个人帐户 As Double, cur起付线 As Currency
    Dim cur公务员补助 As Double, cur医疗照顾公务员补助 As Double, cur超大额公务员补助 As Double, int生育标志 As Integer
    Dim dblHIS总费用 As Double, dbl结算总费用 As Double, dbl差额 As Double, dbl医保总费用 As Double
    
    Dim int结算方式 As Integer, str单病种编码 As String
    Dim int清算方式 As Integer, str单病种 As String
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, datCurr As Date, strNO As String
    Dim str就诊顺序号 As String, str结算编号 As String
    Dim str支付类型 As String, str保险类别 As String
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String
    Dim str结算XML As String, str清算方式 As String, str清算病种编码 As String, str特殊结算说明 As String, str特殊结算标志 As String
    
    If mlng病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人没有完成医保的预结算操作，不能进行结算。"
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    '取主页ID
    gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
    If Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34 Then
        gstrSQL = " Select A.医保号,A.保险类别,B.住院次数 AS 主页ID,A.单病种,A.清算方式,A.单病种编码_结算,A.结算方式,C.入院方式,A.支付类型 " & _
              " From 保险帐户 A,病人信息 B,病案主页 C " & _
              " Where A.病人ID=B.病人ID And B.病人ID=C.病人ID And B.主页ID=C.主页ID And A.病人ID=[1]"
    Else
        gstrSQL = " Select A.医保号,A.保险类别,B.住院次数 AS 主页ID,A.单病种,A.清算方式,A.单病种编码_结算,A.结算方式,C.入院方式,A.支付类型 " & _
              " From 保险帐户 A,病人信息 B,病案主页 C " & _
              " Where A.病人ID=B.病人ID And B.病人ID=C.病人ID And B.住院次数=C.主页ID And A.病人ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取主页ID", lng病人ID)
    cur帐户余额 = 个人余额_贵阳(rsTemp!医保号)
    lng主页ID = rsTemp!主页ID
    str保险类别 = Nvl(rsTemp!保险类别)
    str单病种 = Nvl(rsTemp!单病种)
    int清算方式 = Nvl(rsTemp!清算方式, 1)
    str单病种编码 = Nvl(rsTemp!单病种编码_结算)
    int结算方式 = Nvl(rsTemp!结算方式, 0)
    str支付类型 = Nvl(rsTemp!支付类型, "31")
    
    '检查是否存在未设定特殊药品的数据，有则不允许结算
   If mbln特殊药品提示 Then
        gstrSQL = " Select distinct A.ID" & _
                  " From " & IIf(gblnHIS1026 = True, "住院费用记录", "病人费用记录") & " A,收费细目 B" & _
                  " Where Nvl(A.实收金额,0)<>0 And Nvl(A.数次,0)<>0 And Nvl(A.附加标志,0)<>9 And Nvl(A.记录状态,0)<>0 And Nvl(A.婴儿费,0)=0" & _
                  " And B.类别 IN ('5','6','7') And B.说明 Is Not NULL And A.收费细目ID+0=B.ID And A.病人ID=[1] And A.主页ID=[2]" & _
                  " And Nvl(A.是否上传, 0) = 0 And nvl(A.实收金额,0)<>0 " & _
                  " MINUS" & _
                  " Select distinct 费用ID From 特殊药品收费" & _
                  " Where 病人ID=[1] And 主页ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未设定特殊药品的数据，有则不允许结算", lng病人ID, lng主页ID)
        If rsTemp.RecordCount <> 0 Then
            Err.Raise 9000, gstrSysName, "还存在部分特殊药品未审批，不允许办理结算！"
            Exit Function
        End If
    End If
    
    str结算XML = mdomInput.xml
    '取清算方式与清算病种，待会做为结算入参传上去
'    <RECKONINGTYPE>清算方式</RECKONINGTYPE>
'    <SINGLEILLNESSCODE>单病种编码</SINGLEILLNESSCODE>
    str清算方式 = GetElemnetValue("RECKONINGTYPE")
    str清算病种编码 = GetElemnetValue("SINGLEILLNESSCODE")
    str特殊结算标志 = GetElemnetValue("SPECCALFLAG")
    str特殊结算说明 = GetElemnetValue("SPECCALFLAGTXT")
    
    '必须强制刷卡
    mstr卡号 = ""
    mstr密码 = ""
    '获取病人信息
    If Get验证_贵阳(1, str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
    '获取磁卡信息
    str卡号 = frm刷卡.ShowME
    If str卡号 = "" Then Exit Function
    str密码 = Split(str卡号, "|")(1)
    str卡号 = Split(str卡号, "|")(0)
    
    Screen.MousePointer = vbHourglass
    
    '求个人帐户支付金额
    gstrSQL = "Select Nvl(冲预交,0) as 金额 From 病人预交记录 Where 结算方式='个人帐户' And 记录性质 Not In (11,1) And 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "住院结算", lng结帐ID)
    If Not rsTemp.EOF Then cur个人帐户 = rsTemp("金额")
    '求单据号
    gstrSQL = "Select NO,收费时间 From 病人结帐记录 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "住院结算", lng结帐ID)
    
    'XML文档已经完成初始化，此时只需要更新部分值
    Call InitXML
    Call mdomInput.loadXML(str结算XML)
    Call SetElemnetValue("CARDTYPE", gintType)
    Call SetElemnetValue("CARDDATA", IIf(gintType = 3, "", str卡号))
    Call SetElemnetValue("SNO", gstrIDNO)
    Call SetElemnetValue("PSAMNO", gstrPSAMNO)
    Call SetElemnetValue("ISCAL", "1")
    Call SetElemnetValue("ACCTWANTTOPAY", Format(cur个人帐户, "0.00"))
    Call SetElemnetValue("INVOICENO", "Z_" & rsTemp("NO"))
    Call SetElemnetValue("DODATE", Format(rsTemp("收费时间"), "yyyy-MM-dd HH:mm:ss"))
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str密码)     ' 正式结算时传入密码
'    <RECKONINGTYPE>清算方式</RECKONINGTYPE>
'    <SINGLEILLNESSCODE>单病种编码</SINGLEILLNESSCODE>
    Call InsertChild(mdomInput.documentElement, "RECKONINGTYPE", str清算方式) ' 清算方式
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", str清算病种编码) ' 单病种编码
    
    '预算时已经传递，结帐不需要再传递明细数据
    Call SetElemnetValue("ROWSET", "")
    '调用接口
    If CommServer("CALHOSP", IIf(mbln无密码结算, "1", "0")) = False Then Exit Function
    
    cur全自付 = Val(GetElemnetValue("FEEOUT"))
    cur挂钩自付 = Val(GetElemnetValue("FEESELF"))
    cur起付线 = Val(GetElemnetValue("STARTFEE"))
    cur基数自付 = Val(GetElemnetValue("ENTERSTARTFEE"))
    cur统筹支付 = Val(GetElemnetValue("FUND1PAY")) + Val(GetElemnetValue("ALLOWFUND"))
    cur统筹自付 = Val(GetElemnetValue("FUND1SELF"))
    cur大病统筹 = Val(GetElemnetValue("FUND2PAY"))
    cur大病自付 = Val(GetElemnetValue("FUND2SELF"))
    cur超限自付 = Val(GetElemnetValue("FEEOVER"))
    
'    <FUND3PAY>公务员补助支付</FUND3PAY>
'    <CAREPAY>医疗照顾人员特项公务员补助</CAREPAY>
'    <FUND3OVER>超大额限额公务员补助</ FUND3OVER >
'    <BEARINGFLAG>生育标志</BEARINGFLAG>
    dbl医保总费用 = Val(GetElemnetValue("FEEALL"))
    cur公务员补助 = Val(GetElemnetValue("FUND3PAY"))
    dbl结算总费用 = Val(GetElemnetValue("CALFEEALL"))
    dblHIS总费用 = Val(GetElemnetValue("HOSPFEEALL"))
    dbl差额 = dblHIS总费用 - dbl结算总费用
    
    str结算编号 = GetElemnetValue("BALANCEID")
    str就诊顺序号 = GetElemnetValue("BILLNO")
    
    '填写结算表
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    '帐户年度信息
    Call Get帐户信息(TYPE_贵阳市, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_贵阳市 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 + cur统筹支付 + cur统筹自付 + cur基数自付 + cur超限自付 + cur大病统筹 + cur大病自付 & "," & _
        cur统筹报销累计 + cur统筹支付 + cur大病统筹 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_贵阳市 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户余额 & "," & cur帐户余额 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & ",NULL," & cur基数自付 & "," & _
        g结算数据.发生费用金额 & "," & cur全自付 & "," & cur挂钩自付 & "," & _
        cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & "," & cur大病自付 & "," & cur超限自付 & "," & cur个人帐户 & "," & _
        "'" & str结算编号 & "'," & lng主页ID & ",null,'" & str就诊顺序号 & "',0,'" & AnalyseComputer & "','" & gstrVersion & "','" & str支付类型 & "','" & str就诊顺序号 & "'," & _
            "NULL,'" & str单病种编码 & "','" & str保险类别 & "',to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    gstrSQL = "zl_结算附加信息_Insert (" & lng结帐ID & ",0,0,0," & cur公务员补助 & "," & cur超大额公务员补助 & "," & cur医疗照顾公务员补助 & "," & int生育标志 & "," & _
        "'" & str单病种编码 & "'," & int结算方式 & ",'" & str单病种 & "'," & int清算方式 & "," & dbl结算总费用 & "," & dbl医保总费用 & ",'" & gstr工伤认定编号 & "'," & gint工伤康复住院 & ",NULL,'" & str特殊结算标志 & "','" & str特殊结算说明 & "')"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    '保险结算计算
    gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",0," & cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    '如果清算方式不是按日清单且人员类别不是离休人员，并且当前在院的医保病人，则提示操作员为该病人办理出院手续
    gstrSQL = "Select 单病种,人员身份 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取清算方式", TYPE_贵阳市, lng病人ID)
    If Right(rsTemp!单病种, 1) <> 4 And Not (rsTemp!人员身份 = "市属离休" Or rsTemp!人员身份 = "省属离休") And 医保病人已经出院(lng病人ID) = False Then
        MsgBox "请为该参保人员办理出院手续！"
    End If
    
    住院结算_贵阳 = True
    
    '办理医保出院（如果参数不是在HIS出院同时办理医保出院的话，就需要在结算成功后办理医保出院；如果办理失败，可以保险帐户中再次办理医保出院）
    If mbln医保出院 = False And 医保病人已经出院(lng病人ID) Then
        Call 出院登记_贵阳(lng病人ID, lng主页ID, True)
    End If
    
    '病历登记信息更新
    gstrSQL = "Zl_病历登记记录_贵阳_结帐更新(" & TYPE_贵阳市 & "," & lng病人ID & "," & lng主页ID & ",1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_贵阳(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    Dim lng冲销ID As Long, lng病人ID As Long, lng主页ID As Long
    Dim str结帐日期 As String, str当前日期 As String
    Dim rsTemp  As New ADODB.Recordset, rsCheck As New ADODB.Recordset
    
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, str人员类别 As String
    Dim str就诊顺序号 As String, str结算编号 As String
    Dim cur个人帐户 As Currency
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim curDate As Date    '退费
    Dim bln离休 As Boolean
    Dim str支付类型 As String
    
    On Error GoTo errHand
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
        " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", TYPE_贵阳市, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    lng病人ID = rsTemp!病人ID: lng主页ID = Nvl(rsTemp!主页ID)
    cur个人帐户 = Nvl(rsTemp!个人帐户支付, 0)
    str结算编号 = IIf(IsNull(rsTemp!支付顺序号), "", rsTemp!支付顺序号)
    str就诊顺序号 = IIf(IsNull(rsTemp!备注), "", rsTemp!备注)
    str支付类型 = Nvl(rsTemp!医疗类别)
    If str支付类型 = "" Then
        'Modify By 程池富 2010-01-16 从保险帐户中提取支付类型
        gstrSQL = "Select 支付类型 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "取入院方式", TYPE_贵阳市, lng病人ID)
        str支付类型 = Nvl(rsCheck!支付类型, "31")      ' 支付类别 31：住院；37：转院
    End If
'
'    '判断是否为离体人员
'    gstrSQL = "Select 人员身份 From 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_贵阳市
'    Call OpenRecordset(rsCheck, "判断是否为离休人员")
'    If Not (rsCheck!人员身份 = "省属离休" Or rsCheck!人员身份 = "市属离休") Then
'        MsgBox "基本医疗产生的结帐记录不允许冲销！", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    '非本月结帐的单据，不允许冲销
    gstrSQL = "select to_char(收费时间,'yyyy-MM-dd') 结帐时间 From 病人结帐记录 Where ID=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "取结帐日期", lng结帐ID)
    str结帐日期 = Format(rsCheck!结帐时间, "yyyyMM")
    str当前日期 = Format(zlDatabase.Currentdate, "yyyyMM")
    If str当前日期 <> str结帐日期 Then
        Err.Raise 9000, gstrSysName, "只能冲销本月的结帐单据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '----准备冲销结帐----
    '读取医保病人的基本信息
    gstrSQL = "Select 卡号,医保号,顺序号 中心,人员身份,密码 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保病人的基本信息", TYPE_贵阳市, lng病人ID)
    str卡号 = rsCheck!卡号
    str医保号 = rsCheck!医保号
    str人员类别 = rsCheck!人员身份
    str人员类别 = Switch(str人员类别 = "在职", "11", str人员类别 = "退休", "21", _
                    str人员类别 = "省属离休", "32", str人员类别 = "市属离休", "34", _
                    str人员类别 = "普通居民", "41", str人员类别 = "低保对象", "42", _
                    str人员类别 = "三无人员", "43", str人员类别 = "低收家庭", "44", _
                    str人员类别 = "重度残疾", "45", True, "11") '界面上仍显示低收入家庭，但数据库只有8位，只能保存为低收家庭
    str密码 = IIf(IsNull(rsCheck!密码), "", rsCheck!密码)
    bln离休 = (str人员类别 = "32" Or str人员类别 = "34")
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", str就诊顺序号)            ' 就诊顺序号
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str结算编号)           ' 结算编号
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str支付类型)           ' 支付类型
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)           ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    
    '调用接口
    If CommServer("RETBALANCE", IIf(bln离休, 1, 0)) = False Then Exit Function
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng冲销ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    '帐户年度信息
    Call Get帐户信息(TYPE_贵阳市, rsTemp("病人ID"), Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_贵阳市 & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_贵阳市 & "," & rsTemp("病人ID") & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & rsTemp("发生费用金额") * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") & "," & _
        -1 * Nvl(rsTemp!个人帐户支付, 0) & ",'" & Nvl(rsTemp!支付顺序号) & "',null,null,'" & Nvl(rsTemp!备注) & "'," & _
        "0,'" & AnalyseComputer & "','" & gstrVersion & "','" & str支付类型 & "','" & Nvl(rsTemp!就诊流水号) & "'," & _
        "NULL,'" & Nvl(rsTemp!病种名称) & "','" & Nvl(rsTemp!并发症) & "',to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    
    gstrSQL = "Select * From 结算附加信息 Where 结帐ID=" & lng结帐ID
    Call OpenRecordset_OtherBase(rsTemp, "提取结算附加记录", gstrSQL, gcnGYYB)
    If rsTemp.RecordCount <> 0 Then
        gstrSQL = "zl_结算附加信息_Insert (" & lng冲销ID & "," & -1 * Nvl(rsTemp!公务员补助起付标准, 0) & "," & -1 * Nvl(rsTemp!公务员补助起付线, 0) & "," & -1 * Nvl(rsTemp!普通门诊公务员补助累计, 0) & "," _
            & -1 * Nvl(rsTemp!公务员补助, 0) & "," & -1 * Nvl(rsTemp!超大额公务员补助, 0) & "," & -1 * Nvl(rsTemp!医疗照顾人员特项公务员补助, 0) & "," & rsTemp!生育标志 & "," & _
            "'" & Nvl(rsTemp!单病种编码_结算) & "'," & Nvl(rsTemp!结算方式, 0) & ",'" & Nvl(rsTemp!单病种) & "'," & Nvl(rsTemp!清算方式, 1) & "," & -1 * Nvl(rsTemp!结算总费用, 0) & "," & -1 * Nvl(rsTemp!医保总费用, 0) & ",'" & Nvl(rsTemp!工伤认定编号) & "'," & Nvl(rsTemp!工伤康复住院, 0) & ")"
        gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    End If
    '病历登记信息更新
    gstrSQL = "Zl_病历登记记录_贵阳_结帐更新(" & TYPE_贵阳市 & "," & lng病人ID & "," & lng主页ID & ",0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "贵阳医保")
    住院结算冲销_贵阳 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Sub 查询欠费单位_贵阳(ByVal str单位编码 As String, ByVal str保险类别 As String)
'功能：调用接口查询欠费单位
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str提示 As String
    
    If str单位编码 = "" Then Exit Sub
'    str单位编码 = String(12 - Len(str单位编码), "0") & str单位编码
    
    On Error GoTo errHandle
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "DEPTCODE", str单位编码)                '单位编码
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", str保险类别)              '保险类别
    
    '调用接口
    If CommServer("QUERYARREARDEPT") = False Then Exit Sub
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then
        MsgBox "病人单位无欠费情况。", vbInformation, gstrSysName
        Exit Sub
    End If
    '根据编码得到险种名称
    For Each nodRow In nodRowset.childNodes
        Select Case GetAttributeValue(nodRow, "INSUREKIND")
            Case "3"
                str提示 = str提示 & "、基本医疗"
            Case "8"
                str提示 = str提示 & "、大额医疗"
            Case "10"
                str提示 = str提示 & "、公务员补助"
        End Select
    Next
    
    If str提示 <> "" Then
        MsgBox "病人单位以下险种有欠费情况：" & Mid(str提示, 2) & "。", vbInformation, gstrSysName
    Else
        MsgBox "病人单位无欠费情况。", vbInformation, gstrSysName
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function 错误信息_贵阳(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息

End Function

Public Function 医保项目_贵阳(rsTemp As ADODB.Recordset, Optional ByVal str类别 As String = "12") As Boolean
'功能：医保诊疗药品目录查询
'以前按中心查询，现改为按项目支付类别查询（41-工伤门诊 42-工伤住院 21-离休门诊 22-离休住院 11-普通门诊 12-普通住院 31-生育门诊 32-生育住院）
'按普通住院下载项目清单供对照，对照模式与以前一样，只是提供个查询的界面，可按用户要求查询某个类别下的项目的支付比例等信息
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str编码 As String, str名称 As String, str简码, str备注 As String
    Dim str开始日期 As String, str结束日期 As String, str当前日期 As String
        
    On Error GoTo errHandle
    
    If 医保初始化_贵阳 = False Then Exit Function
    
    str当前日期 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "ITEMCODE", "")         ' 医保编码
    Call InsertChild(mdomInput.documentElement, "ITEMPAYTYPE", str类别) ' 项目支付类别
    
    '调用接口
    If CommServer("QUERYSERVICE") = False Then Exit Function
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then Exit Function
    For Each nodRow In nodRowset.childNodes
        str编码 = GetAttributeValue(nodRow, "ITEMCODE")
        str名称 = ToVarchar(Replace(GetAttributeValue(nodRow, "ITEMNAME"), "'", ""), 40)
        str简码 = ToVarchar(zlCommFun.SpellCode(str名称), 10)
        str开始日期 = Mid(GetAttributeValue(nodRow, "STARTDATE"), 1, 10)
        str结束日期 = Mid(GetAttributeValue(nodRow, "ENDDATE"), 1, 10)
'        PRICELMT           '最高限价
'        SELFRATE           '自付比例
'        BEARINGITEMFLAG    '生育项目标志
'        GSITEMFLAG         '工伤项目标志
'        SPECPAYFLAG        '特殊报销项目标志
'        BGITEMTYPE         '包干结算项目类别
        str备注 = Val(GetAttributeValue(nodRow, "PRICELMT")) & "|" & Val(GetAttributeValue(nodRow, "SELFRATE")) & "|" & _
                  Val(GetAttributeValue(nodRow, "BEARINGITEMFLAG")) & "|" & Val(GetAttributeValue(nodRow, "GSITEMFLAG")) & "|" & _
                  Val(GetAttributeValue(nodRow, "SPECPAYFLAG")) & "|" & Val(GetAttributeValue(nodRow, "BGITEMTYPE"))
        
        If str编码 <> "" And str当前日期 >= str开始日期 And str当前日期 <= str结束日期 Then
            rsTemp.AddNew Array("CLASSCODE", "CODE", "NAME", "PY", "MEMO"), Array("1", str编码, str名称, str简码, str备注)
            rsTemp.Update
        End If
    Next
    
    医保项目_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Public Function InitXML() As Boolean
'功能：初始化XML，增加声明和根节点
    Dim pi As MSXML2.IXMLDOMProcessingInstruction
    Dim nodData As MSXML2.IXMLDOMElement
    
    On Error Resume Next
    
    Set mdomInput = New MSXML2.DOMDocument
    Set mdomOutput = New MSXML2.DOMDocument
    If Err <> 0 Then
        Err.Clear
        Exit Function
    End If
    
'    'XML声明
'    Set pi = mdomInput.createProcessingInstruction("xml", "version=""1.0"" encoding=""GB2312"" standalone=""yes""")
'    mdomInput.appendChild pi
    
    '根节点
    Set nodData = mdomInput.createElement("DATA")
    Set mdomInput.documentElement = nodData
    
    InitXML = True
End Function

Public Function InsertChild(nodParent As MSXML2.IXMLDOMElement, ByVal Name As String, ByVal Value As String) As MSXML2.IXMLDOMElement
'功能：在指定XML元素下增加子元素
    Set InsertChild = mdomInput.createElement(Name)
    InsertChild.Text = Value
    
    nodParent.appendChild InsertChild
End Function

Public Sub InsertAttrib(nodParent As MSXML2.IXMLDOMElement, ByVal Name As String, ByVal Value As String)
'功能：在指定XML元素下增加属性
    Dim attTemp As MSXML2.IXMLDOMAttribute
    
    Set attTemp = mdomInput.createAttribute(Name)
    attTemp.Text = Value
    
    nodParent.setAttributeNode attTemp
End Sub

Public Function CommRecServer(ByVal strFunction As String) As Boolean
'功能：调用医保部件
    Dim InvokeServer As String '调用前置服务器的返回值
    Dim StrInput As String
    
    '参数的传入
    StrInput = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?>" & vbCrLf & mdomInput.xml
    Call DebugTool(StrInput)
    
    Select Case strFunction
        Case "APPRECM"
            InvokeServer = obj清算.APPRECM("ZFRJ", StrInput)
        Case "DELRECM"
            InvokeServer = obj清算.DELRECM("ZFRJ", StrInput)
        Case "APPRECB"
            InvokeServer = obj清算.APPRECB("ZFRJ", StrInput)
        Case "DELRECB"
            InvokeServer = obj清算.DELRECB("ZFRJ", StrInput)
        Case "APPRECG"
            InvokeServer = obj清算.APPRECG("ZFRJ", StrInput)
        Case "DELRECG"
            InvokeServer = obj清算.DELRECG("ZFRJ", StrInput)
        Case "QUERYREC"
            InvokeServer = obj清算.QUERYREC("ZFRJ", StrInput)
        Case Else
            ShowMsgbox "可能医保接口发生变化，无法继续执行交易，请与软件提供商联系！"
            Exit Function
    End Select
    
    '断点设置处
    If InvokeServer = "" Then
        '调用失败，返回固定的错误信息
        InvokeServer = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?><DATA><RETCODE>-1</RETCODE><INFO>医保服务器调用失败</INFO></DATA>"
    End If
            
    If mdomOutput.loadXML(InvokeServer) = False Then
        ShowMsgbox "医保服务器返回值格式不正确。"
    Else
        '再对整个调用是否成功进行分析
        If Val(GetElemnetValue("RETCODE")) = 0 Then
            '调用成功
            CommRecServer = True
        Else
            '调用失败
            InvokeServer = GetElemnetValue("INFO")
            If InvokeServer = "" Then InvokeServer = "服务器调用失败。"
            ShowMsgbox "医保服务器返回错误：" & vbCrLf & vbCrLf & InvokeServer
        End If
    End If
End Function

Public Function CommServer(ByVal strFunction As String, Optional ByVal strAdvance As String = "") As Boolean
'功能：调用医保部件
'strOutPut:医保接口返回信息  2011-05-16程池富增加
    Dim InvokeServer        As String '调用前置服务器的返回值
    Dim StrInput            As String
    Dim strDetailLog        As String
    On Error GoTo errHand
    '参数的传入
    StrInput = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?>" & vbCrLf & mdomInput.xml
    
    strDetailLog = ""
    strDetailLog = strDetailLog & vbCrLf & "操作员：" & UserInfo.姓名
    strDetailLog = strDetailLog & vbCrLf & "工作站：" & AnalyseComputer
    
    Select Case strFunction
        Case "GETPSNINFO"
            InvokeServer = gobj贵阳.GETPSNINFO("ZFRJ", StrInput)
        Case "GETGSINFO"
            InvokeServer = gobj贵阳.GETGSINFO("ZFRJ", StrInput)
        Case "MODIFYCARD"               '修改卡密码
            InvokeServer = gobj贵阳.MODIFYCARD("ZFRJ", StrInput)
        Case "GETCLINNO"                '门诊挂号
            
            InvokeServer = gobj贵阳.GETCLINNO("ZFRJ", StrInput)
        Case "CALCLIN"                  '普通门诊支付
            strDetailLog = strDetailLog & vbCrLf & "支付方式：普通门诊支付"
            strDetailLog = strDetailLog & vbCrLf & "功能：GETCLINNO"
            strDetailLog = strDetailLog & vbCrLf & "入参：" & vbCrLf & StrInput
            Call DebugTool(StrInput)
            InvokeServer = gobj贵阳.CALCLIN("ZFRJ", StrInput)
            strDetailLog = strDetailLog & vbCrLf & "出参：" & vbCrLf & InvokeServer
            
        Case "CALSPECCLIN"              '特殊门诊支付
            strDetailLog = strDetailLog & vbCrLf & "支付方式：特殊门诊支付"
            strDetailLog = strDetailLog & vbCrLf & "功能：CALSPECCLIN"
            strDetailLog = strDetailLog & vbCrLf & "入参：" & vbCrLf & StrInput
            Call DebugTool(StrInput)
            InvokeServer = gobj贵阳.CALSPECCLIN("ZFRJ", StrInput)
            strDetailLog = strDetailLog & vbCrLf & "出参：" & vbCrLf & InvokeServer
            
        Case "RETBALANCE"               '退票
            If strAdvance = "1" Then    '离休退票
                InvokeServer = gobj贵阳.RETLX("ZFRJ", StrInput)
            Else
                InvokeServer = gobj贵阳.RETBALANCE("ZFRJ", StrInput)
            End If
        Case "HOSPREG"                  '住院登记
            InvokeServer = gobj贵阳.HOSPREG("ZFRJ", StrInput)
        Case "HOSPOUT"                  '出院登记
            InvokeServer = gobj贵阳.HOSPOUT("ZFRJ", StrInput)
        Case "CALHOSP"                  '住院支付
            If strAdvance = "1" Then    '无卡结算，仅适用于病人逃费的情况
                InvokeServer = gobj贵阳.CALHOSPSP("ZFRJ", StrInput)
            Else
                InvokeServer = gobj贵阳.CALHOSP("ZFRJ", StrInput)
            End If
        Case "SETRECKONINGTYPE"         '设置清算方式
            InvokeServer = gobj贵阳.SETRECKONINGTYPE("ZFRJ", StrInput)
        Case "QUERYHOSPSINGLEILLNESS"   '单病种清算数据
            InvokeServer = gobj贵阳.QUERYHOSPSINGLEILLNESS("ZFRJ", StrInput)
        Case "QUERYHOSPSINGLEILLNESS_BG"   '单病种结算目录
            InvokeServer = gobj贵阳.QUERYHOSPSINGLEILLNESS_BG("ZFRJ", StrInput)
        Case "QUERYSERVICE"              '医保诊疗药品目录查询
            InvokeServer = gobj贵阳.QUERYSERVICE("ZFRJ", StrInput)
        Case "QUERYARREARDEPT"          '查询欠费单位
            InvokeServer = gobj贵阳.QUERYARREARDEPT("ZFRJ", StrInput)
        Case "GETHOSPSINGLEILLNESS"     '下载单病种清算数据
            InvokeServer = gobj贵阳.GETHOSPSINGLEILLNESS("ZFRJ", StrInput)
        Case "GETHOSPSINGLEILLNESS_BG"  '下载单病种结算目录
            InvokeServer = gobj贵阳.GETHOSPSINGLEILLNESS_BG("ZFRJ", StrInput)
        Case "SETBEARINGFLAG"           '设置生育标志
            InvokeServer = gobj贵阳.SETBEARINGFLAG("ZFRJ", StrInput)
        Case "UPLOADICD"                '上传ICD编码
            InvokeServer = gobj贵阳.UPLOADICD("ZFRJ", StrInput)
        Case "SETCALTYPE"
            InvokeServer = gobj贵阳.SETCALTYPE("ZFRJ", StrInput)
        Case "RETHOSPOUT"
            InvokeServer = gobj贵阳.RETHOSPOUT("ZFRJ", StrInput)
        Case "GETSPECILLNESS"
            InvokeServer = gobj贵阳.GETSPECILLNESS("ZFRJ", StrInput)
        Case "QUERYSPECILLNESS"
            InvokeServer = gobj贵阳.QUERYSPECILLNESS("ZFRJ", StrInput)
        Case "UPLOADBYBILLNO"
            InvokeServer = gobj贵阳.UPLOADBYBILLNO("ZFRJ", StrInput)
        Case Else
            ShowMsgbox "可能医保接口发生变化，无法继续执行交易，请与软件提供商联系！"
            Exit Function
    End Select
    
    '断点设置处
    If InvokeServer = "" Then
        '调用失败，返回固定的错误信息
        InvokeServer = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?><DATA><RETCODE>-1</RETCODE><INFO>医保服务器调用失败</INFO></DATA>"
    End If
            
    If mdomOutput.loadXML(InvokeServer) = False Then
        ShowMsgbox "医保服务器返回值格式不正确。"
    Else
        '再对整个调用是否成功进行分析
        If Val(GetElemnetValue("RETCODE")) = 0 Then
            '调用成功
            CommServer = True
        Else
            '调用失败
            InvokeServer = GetElemnetValue("INFO")
            If InvokeServer = "" Then InvokeServer = "服务器调用失败。"
            ShowMsgbox "医保服务器返回错误：" & vbCrLf & vbCrLf & InvokeServer
            
        End If
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function Get保险参数_贵阳(ByVal str参数名 As String) As String
'功能：获得保险参数
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.参数名,A.参数值 from 保险参数 A " & _
              " where A.参数名=[1] and A.险类=[2] and A.中心 is null "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "贵阳医保", str参数名, TYPE_贵阳市)
    
    If rsTemp.EOF = False Then
        Get保险参数_贵阳 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
    End If
End Function

Public Function SetElemnetValue(ByVal Name As String, ByVal Value As String) As Boolean
'功能：得到指定元素的值
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set xmlElement = mdomInput.documentElement.selectSingleNode(Name)
    If Not xmlElement Is Nothing Then
        '找到指定子元素
        xmlElement.nodeTypedValue = Value
        SetElemnetValue = True
    End If
End Function

Public Function GetElemnetValue(ByVal Name As String) As String
'功能：得到指定元素的值
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set xmlElement = mdomOutput.documentElement.selectSingleNode(Name)
    If Not xmlElement Is Nothing Then
        '找到指定子元素
        GetElemnetValue = xmlElement.Text
'    Else
'        '取消
'        Debug.Assert False
    End If
End Function

Public Function GetAttributeValue(xmlElement As MSXML2.IXMLDOMElement, ByVal Name As String) As String
'功能：得到指定属性的值
    Dim varAttribute As Variant
    
    varAttribute = xmlElement.getAttribute(Name)
    If IsNull(varAttribute) = False Then
        GetAttributeValue = varAttribute
    End If
End Function

Public Function Get验证_贵阳(bytType As Byte, str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, _
                ByVal lng病人ID As Long, Optional bln强制刷卡 As Boolean = False) As Boolean
'功能：得到医保病人的基本功的身份验证信息
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    '从数据库中读出已存储的值
    gstrSQL = " select 卡号,医保号,顺序号,密码,NVL(卡类别,0) AS 卡类别,IDNO,PSAM " & _
              " from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "贵阳医保", lng病人ID, TYPE_贵阳市)
    If rsTemp.EOF = False Then
        str卡号 = IIf(IsNull(rsTemp("卡号")), "", rsTemp("卡号"))
        str卡号 = Replace(str卡号, ",", ";")
        str医保号 = IIf(IsNull(rsTemp("医保号")), "", rsTemp("医保号"))
        str分中心编号 = mstr医保中心编码_贵阳
        gintType = rsTemp!卡类别
        gstrIDNO = Nvl(rsTemp!IDNO)
        gstrPSAMNO = Nvl(rsTemp!PSAM)
    End If
    If bln强制刷卡 = False And lng病人ID > 0 Then
        Get验证_贵阳 = True
        Exit Function
    End If
    
    If frmIdentify贵阳.GetIdentify(bytType, str卡号, str医保号, str分中心编号, str密码) = False Then
        Exit Function
    Else
        '刷卡虽然正确，但要检查是否就是当前病人的
        str卡号 = Split(str卡号, "^")(0)
'        If lng病人ID > 0 Then
'            '从数据库中读出已存储的值
'            gstrSQL = "select 卡号,医保号,顺序号 from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & TYPE_贵阳市
'            Call OpenRecordset(rsTemp, "贵阳医保")
'
'            If str卡号 <> Replace(rsTemp("卡号"), ",", ";") Or str医保号 <> IIf(IsNull(rsTemp("医保号")), "", rsTemp("医保号")) Then
'                MsgBox "当前使用的卡与病人不符。", vbInformation, gstrSysName
'                Exit Function
'            End If
'        End If
    
    End If
    
    Get验证_贵阳 = True
End Function

Public Function OwnerUser(ByVal strUserName As String) As Boolean
    Dim RecUser As New ADODB.Recordset
    '判断当前用户是不是所有者
    OwnerUser = True
    With RecUser
        If .State = 1 Then .Close
        .Open "Select Count(*) 所有者 From ZlSystems Where 所有者='" & strUserName & "'", gcnOracle
        
        If Not .EOF Then
            If Not IsNull(!所有者) Then
                If !所有者 = 0 Then OwnerUser = False
            End If
        End If
    End With
End Function

Public Function Subject(ByVal strData As String) As String
    Dim rsSubject As New ADODB.Recordset
    '返回对应的归属科目编码
    gstrSQL = "" & _
             " Select B.编码,B.类别,A.参数值 归属科目编码   " & _
             " From 保险参数 A,收费类别 B " & _
             " Where A.序号>=6 And A.险类=[1] And A.参数名=B.编码 And B.编码=[2]"
    Set rsSubject = zlDatabase.OpenSQLRecord(gstrSQL, "获取对应的归属科目编码", TYPE_贵阳市, strData)
    
    If rsSubject.EOF Then
        Subject = "11"  '无对应项目返回对应的归属科目编码'11',表示其他
    Else
        Subject = rsSubject!归属科目编码
    End If
End Function

Public Function 门诊挂号_贵阳(ByVal lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    
    Dim datCurr As Date
    Dim str结算方式 As String, arr结算方式
    Dim intTotal  As Integer, intStart As Integer
    Dim cur个人帐户 As Currency
    Dim cur医保基金 As Currency, cur大额统筹 As Currency, cur公务员补助 As Currency, cur差额记帐 As Currency
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, str就诊顺序号 As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    datCurr = zlDatabase.Currentdate()
    gstrSQL = " Select 病人ID From " & IIf(gblnHIS1026 = True, "门诊费用记录", "病人费用记录") & " Where 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人ID", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    'If Get验证_贵阳(0, str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    '对XML对象赋值
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstr医保号)     ' 个人编码
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    '调用接口
    If CommServer("GETCLINNO") = False Then Exit Function
    str就诊顺序号 = GetElemnetValue("BILLNO")
    
    gstrSQL = "Select 病人ID,收费细目ID,数次*NVL(付数,1) AS 数量,标准单价 AS 单价,Nvl(实收金额,0) AS 实收金额,保险编码,'  ' AS 摘要" & _
        " From " & IIf(gblnHIS1026 = True, "门诊费用记录", "病人费用记录") & _
        " Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "贵阳医保", lng结帐ID)
    If Not 门诊虚拟结算_贵阳(rsTemp, str结算方式, "") Then Exit Function
    
    '分解各种结算方式
    arr结算方式 = Split(str结算方式, "|")
    intTotal = UBound(arr结算方式)
    For intStart = 0 To intTotal
        Select Case Split(arr结算方式(intStart), ";")(0)
        Case "个人帐户"
            cur个人帐户 = Val(Split(arr结算方式(intStart), ";")(1))
        Case "医保基金"
            cur医保基金 = Val(Split(arr结算方式(intStart), ";")(1))
        Case "大病统筹"
            cur大额统筹 = Val(Split(arr结算方式(intStart), ";")(1))
        Case "医疗补助"
            cur公务员补助 = Val(Split(arr结算方式(intStart), ";")(1))
        Case "差额记帐"
            cur差额记帐 = Val(Split(arr结算方式(intStart), ";")(1))
        End Select
    Next
    
    If Not 门诊结算_贵阳(lng结帐ID, cur个人帐户, "", True, "1|1") Then Exit Function
    
   '需要修正结算结果
    str结算方式 = ""
    If cur个人帐户 <> 0 Then str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户
    If cur医保基金 <> 0 Then str结算方式 = str结算方式 & "||医保基金|" & cur医保基金
    If cur大额统筹 <> 0 Then str结算方式 = str结算方式 & "||大病统筹|" & cur大额统筹
    If cur公务员补助 <> 0 Then str结算方式 = str结算方式 & "||医疗补助|" & cur公务员补助
    If cur差额记帐 <> 0 Then str结算方式 = str结算方式 & "||差额记帐|" & cur差额记帐 & ";0"
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
    End If
    
    门诊挂号_贵阳 = True
    mlng结帐ID = lng结帐ID
    
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 设置结算方式_贵阳(ByVal lng病人ID As Long, ByVal frmParent As Object, Optional ByVal bln住院 As Boolean = False) As String
    Dim lng主页ID As Long
    Dim rsTemp As New ADODB.Recordset
    '返回结算方式与单病种编码
    
    '工伤保险不存在设置结算方式
    gstrSQL = " Select A.保险类别,B.住院次数 From 保险帐户 A,病人信息 B" & _
              " Where A.病人ID=B.病人ID And A.病人ID=[1] And A.险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险类别", lng病人ID, TYPE_贵阳市)
    If rsTemp!保险类别 = "7" Then Exit Function
    If bln住院 Then lng主页ID = Nvl(rsTemp!住院次数, 0)
    
    设置结算方式_贵阳 = frm设置结算方式.ShowSelect(lng病人ID, lng主页ID, TYPE_贵阳市, frmParent)
End Function

Public Function 设置清算方式_贵阳(ByVal lng病人ID As Long, ByVal frmParent As Object, Optional ByVal bln住院 As Boolean = False) As Boolean
    Dim lng主页ID As Long
    Dim rsTemp As New ADODB.Recordset
    
    '工伤保险不存在设置清算方式
    gstrSQL = " Select A.保险类别,B.住院次数 From 保险帐户 A,病人信息 B" & _
              " Where A.病人ID=B.病人ID And A.病人ID=[1] And A.险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险类别", lng病人ID, TYPE_贵阳市)
    If rsTemp!保险类别 = "7" Then Exit Function
    If bln住院 Then lng主页ID = Nvl(rsTemp!住院次数, 0)
    
    设置清算方式_贵阳 = frm设置清算方式.ShowSelect(lng病人ID, lng主页ID, TYPE_贵阳市, frmParent)
End Function

Public Sub 病种选择_贵阳(ByVal lng病人ID As Long)
    Dim lng病种ID As Long
    Dim str病种 As String
    Dim rs病种 As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    '提取该病人以前的病种信息
    gstrSQL = " select B.编码,B.名称 from 保险帐户 A,保险病种 B " & _
              " where A.病人ID=[1] and A.险类=[2] and A.病种ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "住院预算", lng病人ID, TYPE_贵阳市)
    If rsTemp.RecordCount <> 0 Then
        str病种 = "[" & rsTemp!编码 & "]" & rsTemp!名称
    End If
    
    '供住院病人选择病种
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & TYPE_贵阳市
    Set rs病种 = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种——" & IIf(str病种 = "", "无", str病种))
    If Not rs病种 Is Nothing Then
        lng病种ID = rs病种("ID")
    End If
    
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'病种ID','''" & lng病种ID & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存病种")
End Sub

Public Function 设置ICD编码_贵阳(ByVal lng病人ID As Long) As Boolean
    Dim strICD As String
    Dim rsTemp As New ADODB.Recordset
'    <BILLNO>就诊顺序号</BILLNO>
'    <ICD>ICD编码</ICD>
'    <DODATE>办理时间</DODATE>
    
    If Not 医保病人已经出院(lng病人ID) Then
        MsgBox "该医保病人还未出院！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '选择ICD编码
    strICD = frm病种选择_贵阳.ChooseDisease(lng病人ID)
    If strICD = "" Then Exit Function
    
    '上传病人的ICD编码
    gstrSQL = "Select 医保号,顺序号 From 保险帐户 Where 险类=[1] ANd 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该病人的医保号", TYPE_贵阳市, lng病人ID)
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", Nvl(rsTemp!顺序号))   '顺序号
    Call InsertChild(mdomInput.documentElement, "ICD", strICD)                  '编码
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) '办理日期
    If CommServer("UPLOADICD") = False Then Exit Function
    
    设置ICD编码_贵阳 = True
End Function

Public Function GetItemInsure_贵阳(lng病人ID As Long, lng收费细目ID As Long, bln门诊 As Boolean) As String
    '医保对照类别中插入一条记录
    'insert into 医保对照类别
    '(险类,编码,名称,说明)
    'Values
    '(50,'1','基本','无')
    '将历史对码数据插入到医保对照明细中
    'insert into 医保对照明细
    'select 险类,1,收费细目ID,项目编码,''
    'From 保险支付项目
    'Where 险类 = 50
    Dim strDefault As String            '缺省医保编码
    Dim strCurrent As String            '当前医保编码，门诊取门诊编码，住院取住院编码
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "Select B.类别,A.编码,A.名称,B.说明 From 保险项目 A,医保对照明细 B" & _
        " Where B.险类=[1] And A.险类=B.险类 And A.编码=B.项目编码 And B.收费细目ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保编码", TYPE_贵阳市, lng收费细目ID)
    rsTemp.Filter = "类别=1"
    Select Case rsTemp.RecordCount
    Case 0
        '没有设置对应编码，取缺省编码
        rsTemp.Filter = "类别=0"
        If rsTemp.RecordCount <> 0 Then
            GetItemInsure_贵阳 = rsTemp!编码
        End If
    Case 1
        GetItemInsure_贵阳 = rsTemp!编码
    Case Else
        '多选
        GetItemInsure_贵阳 = frm医保项目选择.ShowSelect(rsTemp, lng收费细目ID)
    End Select
    
    rsTemp.Filter = 0
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    rsTemp.Filter = 0
End Function

Public Function IS离休(ByVal lng病人ID As Long) As Boolean
    Dim str人员身份 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = " Select 人员身份 From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取人员身份", lng病人ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    str人员身份 = Nvl(rsTemp!人员身份)
    str人员身份 = Switch(str人员身份 = "在职", "11", str人员身份 = "退休", "21" _
                  , str人员身份 = "省属离休", "32", str人员身份 = "市属离休", "34", True, "11")
    IS离休 = (str人员身份 = "32" Or str人员身份 = "34")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetNextID(ByVal strTable As String, ByVal cnCustom As ADODB.Connection) As Long
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select " & strTable & "_ID.Nextval From Dual"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, cnCustom
    GetNextID = rsTemp.Fields(0).Value
End Function

Public Sub 贵阳特殊药品提示(ByVal lngItemID As Long)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
'20100607周玉强在此处增加了费用类型为‘药品全自费’的项目时，也提示
    If Not mbln特殊药品提示 Then Exit Sub
    gstrSQL = " Select 类别,名称,说明,nvl(费用类型,'药品全自费') as  费用类型 From 收费细目 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取说明,费用类型字段", lngItemID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    If InStr(1, "5,6,7", rsTemp!类别) = 0 Then Exit Sub
    
    If Nvl(rsTemp!说明) = "" And rsTemp!费用类型 <> "药品全自费" Then Exit Sub
    If Nvl(rsTemp!说明) <> "" Then
    MsgBox rsTemp!名称 + rsTemp!说明, vbInformation, gstrSysName
    End If
    If rsTemp!费用类型 = "药品全自费" Then
    MsgBox rsTemp!名称 + rsTemp!费用类型, vbInformation, gstrSysName
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SaveBalanceLog(ByVal int性质 As Integer, ByVal lng结帐ID As Long, ByVal lng病人ID As Long, _
    ByVal str就诊顺序号 As String, ByVal str结算编号 As String, ByVal str支付类别 As String)
    Dim cnLog As New ADODB.Connection
    Set cnLog = GetNewConnection
    With cnLog
        
        gstrSQL = "ZL_结算日志_贵阳_INSERT(" & lng结帐ID & "," & lng病人ID & ",'" & str就诊顺序号 & "','" & str结算编号 & "'," & _
                "'" & str支付类别 & "','" & gstrUserName & "','" & AnalyseComputer & "',1)"
        .Execute gstrSQL, , adCmdStoredProc
        .Close
    End With
End Sub
 
Private Function logID() As String
    gstrSQL = "select decode(max(日志编码),null,to_char(sysdate,'yyyymmdd') || '000001',to_char(sysdate,'yyyymmdd') ||lpad(to_number(substr(max(日志编码),9,6))+1,6,'0')) from 医保操作日志 where 日志编码 like to_char(sysdate,'yyyymmdd') || '%'"
    logID = zlDatabase.OpenSQLRecord(gstrSQL, "取日志编码").Fields(0)
End Function
'========================================================================================
'=记录修改日志(模块\功能\类型(修改|删除|审批)\修改内容(字符串)\修改文件日志\主键1\主键2\主键3)
'========================================================================================
'=参数:
'=     vstrBM     模块(40)
'=     vstrGN     模块(40)
'=     vstrType   模块(1)
'=     vstrTxt    字符串形式的日志内容
'=     vstrFile   文  件形式的日志内容
'=     vstrKey1   主键1
'=     vstrKey2   主键2
'=     vstrKey3   主键3
'=     vstrKey4   主键4
'=     vstrSource 数据表名,多表用|分开,列如:DEF_CUSTID_M|DEF_CUSTID_M
'=     vblnKillFile 是否删掉日志文件(True则删除,False则不删除)
'========================================================================================
'=注意:
'=     1.vstrBM   大模块,vstrGN 功能,例:vstrBM=营销系统,vstrGN=客户订单
'=     2.vstrType 日志类型:1=修改,2=删除,3=审批
'=     3.vstrTxt,vstrFile只能选择其中的一个参数,如果文件名不能空,vstrFile优先
'=     4.vstrKey1是主记录的主键值1,vstrKey2是主记录的主键值2,vstrKey3是主记录的主键值3
'========================================================================================
Function AddLog(vStrBM As String, vstrGN As String, vstrType As LogType, _
                Optional vstrTxt As String, Optional vstrFile As String, _
                Optional vstrKey1 As String, Optional vstrKey2 As String, Optional vstrKey3 As String, _
                Optional vstrSource As String, Optional vstrKey4 As String, _
                Optional ByVal vblnKillFile As Boolean = False) As Boolean

    Dim AdoRsMs         As ADODB.Recordset
    Dim strUser         As String
    Dim strDate         As String
    Dim strWS           As String
    Dim strFilePath     As String
    Dim strSQL          As String
On Error GoTo ErrH
    strWS = AnalyseComputer                 '取得工作站的名称
    strUser = UserInfo.姓名                 '取得用户ID
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")    '取得服务器日期
    
    strSQL = "select * from 医保操作日志 where 0=1"
    Set AdoRsMs = New ADODB.Recordset
    With AdoRsMs
        .CursorLocation = adUseClient
        .Properties("Initial Fetch Size") = 50
        .PageSize = 50
        .Open strSQL, gcnGYYB, adOpenKeyset, adLockOptimistic, adAsyncFetch     '运行可修改游标
    End With
    With AdoRsMs
        .AddNew
        .Fields("日志编码") = logID     '自动编号
        .Fields("模块") = vStrBM
        .Fields("功能") = vstrGN
        .Fields("类型") = vstrType
        .Fields("日期") = strDate
        .Fields("主键1") = vstrKey1
        .Fields("主键2") = vstrKey2
        .Fields("主键3") = vstrKey3
        .Fields("主键4") = vstrKey4
        .Fields("用户") = strUser
        .Fields("工作站") = strWS
        .Fields("数据来源") = vstrSource
        If Trim(vstrFile) <> "" And Len(Trim(Dir(vstrFile))) > 0 Then
            .Fields("日志类型") = "2"
            .Fields("日志描述") = "<长文本>"
            strFilePath = vstrFile
            Call WriteToDB(AdoRsMs.Fields("日志描述"), strFilePath)
            If vblnKillFile Then Kill strFilePath
        Else
            .Fields("日志类型") = "1"
            .Fields("日志描述") = vstrTxt
        End If
        .Update
    End With
    Set AdoRsMs = Nothing
    Exit Function
ErrH:
    Err.Clear
    Resume Next
    Exit Function
End Function

'========================================================================================
'=功    能:从表中读出所有的信息并写入到文件中(进入修改窗口时)
'=参(1) 数:TableName 表名,用分号隔开
'=参(2) 数:TableCondition条件,用分号隔开,跟表对应,同样的可以只传一个条件
'=参(3) 数:vFile文件名,在修改后传入，不传是修改前
'=参(4) 数:vstrTilte标题名,如果传入则是将要删除的记录
'=返 回 值:返回文件路径
'========================================================================================
Public Function EditFormerWriteFileA(ByVal TABLEName As String, Optional ByVal TableCondition As String = "", Optional vFile As String = "", Optional vstrTilte As String) As String
    Dim Sql         As String
    Dim rs          As ADODB.Recordset
    Dim PrefixRs    As ADODB.Recordset
    Dim i           As Integer
    Dim j           As Integer
    Dim strFile     As String
    Dim a           As TextStream
    Dim fs          As FileSystemObject
    Dim FileName    As String
    Dim str()       As String
    Dim strW()      As String
    Dim strV        As Variant
On Error GoTo ErrH

    str = Split(TABLEName, ";")
    strW = Split(TableCondition, ";")
    
    If Trim(vFile) = "" Then
        EditFormerWriteFileA = logID
        FileName = App.Path & "\" & EditFormerWriteFileA & ".betry"
        Set fs = New FileSystemObject
        Set a = fs.CreateTextFile(FileName, True)
        If Trim(vstrTilte) <> "" Then
            a.WriteLine vbCrLf & Trim(vstrTilte)
        Else
            a.WriteLine vbCrLf & "修改前"
        End If
    Else
        FileName = vFile
        Set fs = New FileSystemObject
        Set a = fs.OpenTextFile(FileName, ForAppending)
        If Trim(vstrTilte) <> "" Then
            a.WriteLine vbCrLf & Trim(vstrTilte)
        Else
            a.WriteLine vbCrLf & "修改后"
        End If
    End If
        
    For j = 0 To UBound(str)
        strFile = ""
        Sql = "SELECT Table_Name as TableName,Column_Name as FieldName,Column_Name as FieldNote FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = '" & str(j) & "'"
        Set PrefixRs = zlDatabase.OpenSQLRecord(Sql, "日志")
        With PrefixRs
            While Not .EOF
                strFile = strFile & IIf(Len(Trim(!FieldNote & "")), !FieldNote, !FieldName) & vbTab
                .MoveNext
            Wend
        End With
        Set PrefixRs = Nothing
        a.Write vbCrLf & "表名【" & str(j) & "】" & vbCrLf
        a.Write strFile
        Sql = "select * from " & str(j)
        If j <= UBound(strW) Then i = j
        If Trim(strW(i)) <> "" Then Sql = Sql & " WHERE " & strW(i)
        Set rs = zlDatabase.OpenSQLRecord(Sql, "日志")
        If rs.RecordCount > 0 Then
            strV = rs.GetString(adClipString, , vbTab, vbCrLf, "")
            a.Write vbCrLf & strV
            Set rs = Nothing
        End If
    Next j
    
    a.Close
    EditFormerWriteFileA = FileName
    Exit Function
ErrH:
    Err.Clear
    Resume Next
    EditFormerWriteFileA = ""
    Exit Function
End Function

'========================================================================================
'=存储文件到数据库
'========================================================================================
Private Function WriteToDB(ByRef COL As ADODB.Field, ByVal FileName As String) As Boolean
    Dim mStream As ADODB.Stream
    Dim Lines As String
    Dim NextLine As String
On Error GoTo ErrH
    Open FileName For Input As #1
    Do Until EOF(1)
        Line Input #1, NextLine
        Lines = Lines & NextLine & Chr(13) & Chr(10)
    Loop
    Close #1
    COL.Value = Lines
    Exit Function
ErrH:
    
    Err.Clear
    Exit Function
End Function

Public Function DrugsUsed(ByVal intinsure As Integer, ByVal lng病人ID As Long, ByVal lng药品ID As Long, ByVal dbl处方单量 As Double) As String
    '返回：超量时返回超量信息，否则返回空串
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Zl_药品超量检测_贵阳(" & intinsure & "," & lng病人ID & "," & lng药品ID & "," & _
                dbl处方单量 & "," & gint累计用药量计算标准 & "," & IIf(gblnHIS1026 = True, 1, 0) & ") As 超标信息 From Dual "
    Call OpenRecordset_OtherBase(rsTemp, gstrSQL)
    If Nvl(rsTemp!超标信息, "") <> "" Then
        DrugsUsed = Nvl(rsTemp!超标信息, "") & vbNewLine
    Else
        DrugsUsed = ""
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    DrugsUsed = ""
End Function

Public Function zl_Ip_Address_FromOrc(Optional strDefaultIp_Address As String = "") As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:通过oracle获取的计算机的IP地址
    '入参:strDefaultIp_Address-缺省IP地址
    '出参:
    '返回:返回IP地址
    '编制:刘兴洪
    '日期:2009-01-21 11:08:47
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strIp_Address As String, strSQL As String
    Err = 0: On Error GoTo errHand:
     strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取IP地址")
    If rsTemp.EOF = False Then
        strIp_Address = zlCommFun.Nvl(rsTemp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = strDefaultIp_Address
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    zl_Ip_Address_FromOrc = strIp_Address
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

'功能说明：对交易进行确认或取消
'strBusiness：交易名称，对应于枚举变量 交易Enum
'blnResult：TRUE表示提交成功，FALSE表示发生异常，需要取消医保交易

'医保接口内增加方法：BusinessAffirm，用来确认或取消某个交易，调用流程：
'1 ?HIS处理
'2 ?医保交易成功
'3 ?HIS提交
'4 ?调用BusinessAffirm确认医保交易
'
'需要考虑从第3步开始，如果出现异常，就调用确认交易，传入FALSE表示需要取消医保交易
'第3步以前出现任何异常不在HIS考虑范围内?
'
'本次修改仅要求：门诊结算、门诊结算作废、住院结算、住院结算作废这四个交易处，HIS进行相应修改。

Public Sub BusinessAffirm_贵阳市(ByVal intBusiness As Integer, ByVal blnResult As Boolean, Optional ByVal intinsure As Integer = 0, _
    Optional ByVal strAdvance As String)
    '处理交易后的提示消息
    If blnResult Then
        '交易成功
        Select Case intBusiness
            Case 交易Enum.Busi_RegistSwap '门诊挂号
                Call frm结算信息.ShowME(mlng结帐ID)
            Case 交易Enum.Busi_RegistDelSwap '门诊挂号冲销
            Case 交易Enum.Busi_ClinicSwap '门诊结算
            Case 交易Enum.Busi_ClinicDelSwap '门诊结算冲销
            Case 交易Enum.Busi_ComeInSwap '入院登记
            Case 交易Enum.Busi_ComeInDelSwap '入院登记撤销
            Case 交易Enum.Busi_SettleSwap '住院结算
            Case 交易Enum.Busi_SettleDelSwap '住院结算冲销
        End Select
    Else
        '交易失败
        Select Case intBusiness
            Case 交易Enum.Busi_RegistSwap '门诊挂号
            Case 交易Enum.Busi_RegistDelSwap '门诊挂号冲销
            Case 交易Enum.Busi_ClinicSwap '门诊结算
            Case 交易Enum.Busi_ClinicDelSwap '门诊结算冲销
            Case 交易Enum.Busi_ComeInSwap '入院登记
            Case 交易Enum.Busi_ComeInDelSwap '入院登记撤销
            Case 交易Enum.Busi_SettleSwap '住院结算
            Case 交易Enum.Busi_SettleDelSwap '住院结算冲销
        End Select
    End If
End Sub

