Attribute VB_Name = "mdl黔南"
Option Explicit
Public Enum 业务类型_黔南
    医保初始化 = 0
    获得参保人员信息
    参保资格审查
    病人登记
    获得通过的审批信息
    更新就诊信息
    确定明细项目单价
    录入处方明细
    处方退方
    项目审批结果查询
    预结算
    结算
    冲正交易
    
    收费目录下载预处理
    收费目录下载处理
    
    读取卡内数据
    读卡组件_扩展信息
End Enum
Private gInitCard As Boolean                '初始化了卡的
Private Type InitbaseInfor
    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    医院编码 As String                      '初始医院编码
    
    解析卡内数据 As Boolean
End Type
Public InitInfor_黔南 As InitbaseInfor

Private Type 病人身份
        卡号        As String
        医保证号    As String       '即医保号
        姓名     As String
        性别     As String
        身份证号 As String
        出生日期  As String
        年龄        As Integer
        类别编码    As String   '人员类别编码
        类别名称    As String   '人员类别名称
        人员状态    As String
        单位编码    As String
        单位名称    As String
        医疗类别    As String '医疗人员类别或就诊类别
        病种编码    As String
        病种名称    As String
        统筹区号    As String
        病种ID      As Long
        审批类别    As String
        帐户余额    As Double
        门诊号      As String
        住院号      As String
        
        审批编号    As String
        项目编码    As String
        项目名称    As String
        
        费用总额    As Double
        动态信息    As String   '所有16位动态信息
        扩展信息    As String
        病人ID      As Long
        
        入院诊断编码 As String
        入院诊断名称    As String
        
        确诊诊断编码    As String
        确诊诊断名称    As String
        中途结算    As Boolean
        下个人帐户 As Boolean
        
End Type

Public g病人身份_黔南 As 病人身份
Public gcnOracle_黔南 As ADODB.Connection     '中间库连接
Private gcnOracle_中心 As ADODB.Connection     '黔南医保中心联接串

Private Type 结算数据
        费用总额    As Double
        统筹支付    As Double
        账户支付    As Double
        现金支付    As Double
        大病垫付    As Double
        
        动态信息    As String
        交易流水号  As String
End Type
Private 虚拟结算数据 As 结算数据

'1 接口初始化
Private Declare Function Init Lib "SiInterface" Alias "INIT" () As Long
'2 业务处理：调用执行医保业务所需要的处理
Private Declare Function BUSINESS_HANDLE Lib "SiInterface" _
    (ByVal StrInput As String, ByVal strOutput As String) As Long

'3 业务查询：调用执行医保业务所需要的处理
Private Declare Function QUERY_HANDLE Lib "SiInterface" _
    (ByVal StrInput As String, ByVal strOutput As String) As Long
    
'4.读写卡组件对象声明
Public gobj黔南 As Object
Public gobj黔南Err As Object
Public gobjTest As Object
Private Const STR_中心维护代码 = "1"
Private mblnInit As Boolean         '是否初始化

Public Function 医保初始化_黔南() As Boolean
    Dim strReg As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    Dim bln读卡器 As Boolean
    
    If mblnInit = True Then
        医保初始化_黔南 = True
        Exit Function
    End If
    
    
    GetRegInFor g公共全局, "医保", "读卡器", strReg
    bln读卡器 = Val(strReg) = 1

    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_黔南.模拟数据 = True
    Else
        InitInfor_黔南.模拟数据 = False
    End If
    
    Call GetRegInFor(g公共模块, "操作", "解析卡内数据", strReg)
    If Val(strReg) = 1 Then
        InitInfor_黔南.解析卡内数据 = True
    Else
        InitInfor_黔南.解析卡内数据 = False
    End If
    
    InitInfor_黔南.解析卡内数据 = InitInfor_黔南.解析卡内数据 Or InitInfor_黔南.模拟数据
    
    '创建黔南医保对象
'    If gInitCard = True And bln读卡器 Then
'        Call sCard_CloseCardWithoutSave
'    End If
    
    If gInitCard = False Then
        Set gobj黔南 = Nothing
        
        Err = 0
        On Error Resume Next
        Set gobj黔南 = CreateObject("SiCard.SiCardOperator")
        If Err <> 0 Then
            If InitInfor_黔南.模拟数据 Then
            Else
                ShowMsgbox "创建东软读写卡组失败!"
                Exit Function
            End If
        End If
        Set gobj黔南Err = CreateObject("SiCommTool.SiErrorCtl")
        If Err <> 0 Then
            If InitInfor_黔南.模拟数据 Then
            Else
                ShowMsgbox "创建东软读写卡组失败!"
                Exit Function
            End If
        End If
         
        '初始化读写卡组件
        If bln读卡器 Then
            If sCard_InitCard = False Then
                If Not InitInfor_黔南.模拟数据 Then
                    Exit Function
                End If
            End If
        End If
        
        '取医院编码
        gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_黔南)
        InitInfor_黔南.医院编码 = Nvl(rsTemp!医院编码)
        
        If Open中间库 = False Then Exit Function
        '初始化医保接口
        If 业务请求_黔南(医保初始化, "", strOutput) = False Then
            Exit Function
        End If
        
        gInitCard = True
    End If
    mblnInit = True
    医保初始化_黔南 = True
End Function
Private Function Open中间库() As Boolean
    '连接中间库
    '中间库连接
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strServer As String, strPass As String
    
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "渝北医保", TYPE_黔南)
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
    Set gcnOracle_黔南 = New ADODB.Connection

    If OraDataOpen(gcnOracle_黔南, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If


    '中间库连接
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '中心%' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "渝北医保", TYPE_黔南)
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "中心用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "中心服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "中心用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_中心 = New ADODB.Connection

    If OraDataOpen(gcnOracle_中心, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中心数据库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    Open中间库 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function 解析卡_黔南() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:解析卡内数据
    '--入参数:strCardData-卡内数据
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strOutput As String
    Dim StrInput As String
    Dim strArr As Variant
    Err = 0
    On Error GoTo errHand:
    
    If InitInfor_黔南.解析卡内数据 Then
        Read模拟数据 读取卡内数据, StrInput, strOutput
        If strOutput = "" Then
            解析卡_黔南 = False
            Exit Function
        End If
        strArr = Split(strOutput, "|")
        With g病人身份_黔南
            .医保证号 = strArr(1)
            .身份证号 = strArr(2)
            .单位编码 = strArr(3)
            .卡号 = strArr(4)
            .姓名 = strArr(5)
            .性别 = strArr(6)
            .出生日期 = strArr(7)
            .类别编码 = strArr(8)
            .人员状态 = strArr(10)
            .帐户余额 = Val(strArr(12))
        End With
        
    Else
        '读取卡
        If sCard_ReadCard = False Then Exit Function
        '获取卡信息
        'bytType:1-SiCardBaseInfo社保卡基本信息
        '        2-SiCardDynaInfo社保卡动态信息
        '        3-SiCardAcctInfo社保卡帐户信息
        '        4-SiCardExtInfo社保卡扩展信息
        If sCard_属性值(1, strOutput) = False Then Exit Function
        '个人编号|身份证号|单位编码|社保卡号|姓名|性别|出身日期|人员类别|参保日期|人员状态|变更日期
        strArr = Split("0|" & strOutput, "|")
        With g病人身份_黔南
            .医保证号 = strArr(1)
            .身份证号 = strArr(2)
            .单位编码 = strArr(3)
            .卡号 = strArr(4)
            .姓名 = strArr(5)
            .性别 = strArr(6)
            .出生日期 = strArr(7)
            .类别编码 = strArr(8)
            .人员状态 = strArr(9)
        End With
        If sCard_属性值(3, strOutput) = False Then Exit Function
        '读取个人帐户余额
        '帐户余额
        strArr = Split(strOutput, "|")
        With g病人身份_黔南
            .帐户余额 = Val(strArr(0))
        End With
        
        '获取动态信息
        If sCard_属性值(2, strOutput) = False Then Exit Function
        With g病人身份_黔南
            .动态信息 = strOutput
        End With
        '获取扩展信息
        If sCard_属性值(4, strOutput) = False Then Exit Function
        With g病人身份_黔南
            .扩展信息 = strOutput
        End With
    End If
    解析卡_黔南 = True
    Exit Function
errHand:
    解析卡_黔南 = False
    ShowMsgbox "IC卡错误,不能识别!"
End Function
Private Function sCard_InitCard() As Boolean
    Dim lngReturn As Long
    Dim strErrInfor As String
    
    '成功返回零 失败返回大于零的错误号
    Err = 0
    On Error GoTo errHand:
    lngReturn = gobj黔南.InitCard()
    If lngReturn <> 0 Then
        '
        strErrInfor = sCard_ErrInfor(lngReturn)
        If strErrInfor <> "" Then
            ShowMsgbox strErrInfor
        End If
        Exit Function
    End If
    sCard_InitCard = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function sCard_属性值(ByVal bytType As Long, strPropertyValue As String, Optional blnWrite As Boolean = False) As Boolean
    Dim lngReturn As Long
    Dim strReturn As String
    Dim strErrInfor As String
    'bytType:1-SiCardBaseInfo社保卡基本信息
    '        2-SiCardDynaInfo社保卡动态信息
    '        3-SiCardAcctInfo社保卡帐户信息
    '        4-SiCardExtInfo社保卡扩展信息
    
    '成功返回零 失败返回大于零的错误号
    sCard_属性值 = False
    
    
    Err = 0
    On Error GoTo errHand:
    Select Case bytType
        Case 1  ' SiCardBaseInfo社保卡基本信息
            '个人编号|身份证号|单位编码|社保卡号|姓名|性别|出身日期|人员类别|参保日期|人员状态|变更日期
            If blnWrite Then
                gobj黔南.SiCardBaseInfo = strPropertyValue
                DebugTool "写动基本信息：" & strPropertyValue
            Else
                strReturn = gobj黔南.SiCardBaseInfo
            End If
            
            
        Case 2  ' SiCardDynaInfo社保卡动态信息
            If blnWrite Then
                gobj黔南.SiCardDynaInfo = strPropertyValue
                DebugTool "写动动态信息：" & strPropertyValue
            Else
                strReturn = gobj黔南.SiCardDynaInfo
            End If
        Case 3  ' SiCardAcctInfo社保卡帐户信息
            '卡帐户余额
            If blnWrite Then
                gobj黔南.SiCardAcctInfo = strPropertyValue
                DebugTool "写帐户余额：" & strPropertyValue
            Else
                strReturn = gobj黔南.SiCardAcctInfo
            End If
        Case Else 'SiCardExtInfo社保卡扩展信息
            '定点医院1|定点医院2|定点医院3|定点医院4|定点医院5|出院日期|住院状态（1—住院；2—出院）|就诊医院|医疗类别
            If blnWrite Then
                gobj黔南.SiCardExtInfo = strPropertyValue
                DebugTool "写扩展信息：" & strPropertyValue
            Else
                strReturn = gobj黔南.SiCardExtInfo
            End If
    End Select
    If blnWrite Then
    Else
        strPropertyValue = strReturn
    End If
    sCard_属性值 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function sCard_CloseCardWithoutSave() As Boolean
    Dim lngReturn As Long
    Dim strErrInfor As String
    
    '读取写组件初始化.
    '初始华医保卡
    '成功返回零 失败返回大于零的错误号
    Err = 0
    On Error GoTo errHand:
    If gobj黔南 Is Nothing Then Exit Function
    lngReturn = gobj黔南.CloseCardWithoutSave()
    If lngReturn <> 0 Then
        '
        strErrInfor = sCard_ErrInfor(lngReturn)
        If strErrInfor <> "" Then
            ShowMsgbox strErrInfor
        End If
        Exit Function
    End If
    sCard_CloseCardWithoutSave = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function sCard_ReadCard() As Boolean
    Dim lngReturn As Long
    Dim strErrInfor As String
    
    '读取写组件初始化.
    '初始华医保卡
    '成功返回零 失败返回大于零的错误号
    Err = 0
    On Error GoTo errHand:
    
    lngReturn = gobj黔南.ReadCard()
    If lngReturn <> 0 Then
        '
        strErrInfor = sCard_ErrInfor(lngReturn)
        If strErrInfor <> "" Then
            ShowMsgbox strErrInfor
        End If
        Exit Function
    End If
    sCard_ReadCard = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function sCard_SaveCard() As Boolean
    Dim lngReturn As Long
    Dim strErrInfor As String
    
    'IDL声明 HRESULT SaveCard([in] BSTR prm_ctlStr, [out, retval] int *ret_appCode)
    '方法声明 int SaveCard (BSTR prm_ctlStr)
    '返 回 值 成功返回零 失败返回大于零的错误号
    '参 数 prm_ctlStr输入 写卡控制串，该控制串决定了组件将以何种方式写卡（全写或者部分写），例如“Athene”代表更新动态信息并忽略其他项信息，“Apollo”表示更新动态和扩展信息并忽略其他项信息。对于非社保中心客户程序而言，仅关心动态信息、帐户信息和扩展信息，所以大多数情况下使用“Apollo”是最合适的
    '说 明 写卡操作是根据读卡时填充进SiCard缓存的信息和用户程序对SiCardDynaInfo等属性进行设置后更改的缓存信息写入卡上。所以在执行写卡操作之前，必须确认希望写卡的信息已经正确的设置给了对应的组件属性
    '卡 类 型 全部支持，其中明华Memory卡支持特定情况不读卡就写卡
    
    Err = 0
    On Error GoTo errHand:
    
    lngReturn = gobj黔南.SaveCard("Apollo")
    If lngReturn <> 0 Then
        strErrInfor = sCard_ErrInfor(lngReturn)
        If strErrInfor <> "" Then
            ShowMsgbox strErrInfor
        End If
        Exit Function
    End If
    DebugTool "写卡成功"
    sCard_SaveCard = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function sCard_SetupCardOption_黔南() As Boolean
    Dim lngReturn As Long
    Dim strErrInfor As String
    
    '读取写组件初始化.
    '初始华医保卡
    '成功返回零 失败返回大于零的错误号
    Err = 0
    On Error GoTo errHand:
    '先初始化卡
    If gobj黔南 Is Nothing Then
         Set gobj黔南 = CreateObject("SiCard.SiCardOperator")
    End If
    lngReturn = gobj黔南.SetupCardOption()
    If lngReturn <> 0 Then
        '
        strErrInfor = sCard_ErrInfor(lngReturn)
        If strErrInfor <> "" Then
            ShowMsgbox strErrInfor
        End If
        Exit Function
    End If
    sCard_SetupCardOption_黔南 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function sCard_ErrInfor(lngReturn As Long) As String
    '获取读写卡组件的错误描述
    Dim strReturn As String
    
    '初始华医保卡
    '成功返回零 失败返回大于零的错误号
    Err = 0
    On Error GoTo errHand:
    If gobj黔南Err Is Nothing Then
        Set gobj黔南Err = CreateObject("SiCommTool.SiErrorCtl")
    End If
    strReturn = gobj黔南Err.Describe(lngReturn)
    sCard_ErrInfor = strReturn
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 获取住院状态_黔南(ByRef lng状态 As Long) As Boolean
    '返回lng状态,(0-门诊,1在院)
    Dim strOutput As String
    Dim strArr As Variant
    
    获取住院状态_黔南 = False
    Err = 0
    On Error GoTo errHand:
    '??lng状态 = GetHospstatus()
    '        1-SiCardBaseInfo社保卡基本信息
    '        2-SiCardDynaInfo社保卡动态信息
    '        3-SiCardAcctInfo社保卡帐户信息
    '        4-SiCardExtInfo社保卡扩展信息
    DebugTool "进入获取住院状态"
    
    If InitInfor_黔南.模拟数据 Then
        '定点医院1|定点医院2|定点医院3|定点医院4|定点医院5|出院日期|住院状态（1—住院；2—出院）|就诊医院|医疗类别
        Call Read模拟数据(读卡组件_扩展信息, "", strOutput)
        If strOutput = "" Then Exit Function
    Else
        If sCard_属性值(4, strOutput) = False Then Exit Function
        '定点医院1|定点医院2|定点医院3|定点医院4|定点医院5|出院日期|住院状态（1—住院；2—出院）|就诊医院|医疗类别
    End If
    strArr = Split("|" & strOutput, "|")
    lng状态 = IIf(Val(strArr(7)) = 1, 1, 0)
    DebugTool "获取住院状态,返回值为:" & lng状态 & " 返回串为:" & strOutput
    获取住院状态_黔南 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 医保终止_黔南() As Boolean
    '结束读写卡组件
    Dim bln读卡器 As Boolean
    Dim strReg As String
    mblnInit = False
    GetRegInFor g公共全局, "医保", "读卡器", strReg
    bln读卡器 = Val(strReg) = 1
    
    If bln读卡器 Then
        If sCard_CloseCardWithoutSave = False Then
            If Not InitInfor_黔南.模拟数据 Then
                Exit Function
            End If
        End If
    End If
    gInitCard = False
    
    Err = 0
    On Error Resume Next
    
    Set gobj黔南 = Nothing
    Set gobj黔南Err = Nothing
    If gcnOracle_黔南.State = 1 Then
        gcnOracle_黔南.Close
    End If
    If Not gcnOracle_中心 Is Nothing Then
        gcnOracle_中心.Close
    End If
    医保终止_黔南 = True
End Function

Public Function 身份标识_黔南(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '返回：空或信息串
    Dim bln读卡器 As Boolean
    Dim strReg As String
    
    GetRegInFor g公共全局, "医保", "读卡器", strReg
    bln读卡器 = Val(strReg) = 1
    If bln读卡器 = False Then
        ShowMsgbox "没有读卡器，不能进行身份验证,请在保险类别中设置"
        Exit Function
    End If
    Err = 0
    On Error GoTo errHand:
    身份标识_黔南 = frmIdentify黔南.GetPatient(bytType, lng病人ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_黔南 = ""
End Function


Public Function 个人余额_黔南(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取个人帐户余额", lng病人ID, TYPE_黔南)
    
    If rsTemp.EOF Then
        个人余额_黔南 = 0
    Else
        个人余额_黔南 = rsTemp("帐户余额")
    End If
End Function
Public Function 参保资格审查_黔南() As Boolean
        '功能：验证当前医保人员的资格审查
        '返回：返回true,否则返回False
        
        Dim StrInput As String
        Dim strOutput As String
        参保资格审查_黔南 = False
        Dim strArr
        
        With g病人身份_黔南
            StrInput = .医保证号 & "|"
            StrInput = StrInput & .卡号 & "|"
            StrInput = StrInput & .统筹区号
        End With
        
        '入参: 医保证号码|IC卡号|统筹区号
        '出参: 封锁级别|封锁原因汉字描述
        '封锁级别:
        '        = 100  没有被封锁
        '        = 1    统筹封锁,属个人封锁
        '        < 0    全封锁
        
        Err = 0
        On Error GoTo errHand:
        If 业务请求_黔南(参保资格审查, StrInput, strOutput) = False Then
            Exit Function
        End If
        
        strArr = Split(strOutput, "|")
        Select Case Val(strArr(1))
            Case 100   '没有被封锁
            Case 1      '统筹封锁,属个人封锁,可以看普通门诊，可以消费卡帐户，但是不能有统筹支付也就是不能使用住院和慢性病等等。
                If g病人身份_黔南.医疗类别 <> "11" Then
                    ShowMsgbox "已经被统筹封锁(属个人封锁)，只能看普能门诊!"
                    Exit Function
                End If
            Case Is < 0 '全封锁
                ShowMsgbox "已经被统筹封锁，不能进行医保结算!"
                Exit Function
            Case Else
                ShowMsgbox strArr(2) & "，不能进行医保结算!"
                Exit Function
        End Select
        参保资格审查_黔南 = True
        Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function

Private Function 门诊明细写入(ByVal rs明细 As ADODB.Recordset, Optional ByVal bln虚拟 As Boolean = True) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr
    Dim strDate As Date
    
    Dim str审批编号 As String
    门诊明细写入 = False
    g病人身份_黔南.费用总额 = 0
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Err = 0
    On Error GoTo errHand:
    '然后插入处方明细
    With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        
        Do Until rs明细.EOF
            gstrSQL = "select A.名称,A.编码,A.类别,A.规格,A.计算单位,B.项目编码,B.附注,B.是否医保,A.计算单位,E.规格,G.名称 剂型 " & _
                      "from 收费细目 A,(Select * From 保险支付项目 where 险类=" & TYPE_黔南 & ") B,药品目录 E ,药品信息 F,药品剂型 G " & _
                      "where A.ID=[1] and A.ID=B.收费细目ID(+) " & _
                     "        AND A.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) "
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", CLng(rs明细("收费细目ID")))
            If rsTemp.EOF = True Then
                MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
                Exit Function
            End If
            
            gstrSQL = "" & _
                  "   Select * From 医保收费目录 " & _
                  "   Where 类别=[1] and 编码=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取医保相关信息", CStr(Nvl(rsTemp!附注)), CStr(Nvl(rsTemp!项目编码)))
            
              
            If Val(Nvl(rs明细("实收金额"), 0)) <> 0 Then
                '住院(门诊)号|处方号|审批编号|医院内码|医保编码|项目名称|费用等级|费用类别|单价|数量|金额|单位|规格|剂型|开方日期|录入标志
                '获取经过审批的项目编号
                If Nvl(rsTemp!是否医保, 0) = 0 Then
                    StrInput = g病人身份_黔南.医保证号 & "|"
                    If Val(Nvl(rsTemp!附注)) = 1 Then
                        '吕说:药品只能传医院编码,而诊疗只能等只能传医保编码
                        StrInput = StrInput & Nvl(rsTemp!编码, "9000900099")
                    Else
                        StrInput = StrInput & Nvl(rsTemp!项目编码, "9000900099")
                    End If
                    
                    If 业务请求_黔南(项目审批结果查询, StrInput, strOutput) = False Then
                        strOutput = "|"
                    End If
                    strArr = Split(strOutput, "|")
                    str审批编号 = strArr(1)
                Else
                    str审批编号 = ""
                End If
                
                StrInput = g病人身份_黔南.门诊号 & "|"
                StrInput = StrInput & g病人身份_黔南.门诊号 & "|"
                StrInput = StrInput & str审批编号 & "|"
                StrInput = StrInput & Nvl(rsTemp!编码) & "|"
                
                StrInput = StrInput & Nvl(rsTemp!项目编码, "9000900099") & "|"
                StrInput = StrInput & Nvl(rsTemp!名称) & "|"
                
                
                If rsTmp.EOF Then
                    '吕说:中草药的需传1
                    If InStr(1, "7中草药", Nvl(!收费类别)) <> 0 Then
                        StrInput = StrInput & "1" & "|"
                    Else
                        StrInput = StrInput & "3" & "|"
                    End If
                    StrInput = StrInput & Split(Get费用类别(Nvl(!收费类别)), "-")(0) & "|"
                Else
                    '刘兴宏:2007/09/28加入,自费部分又是中药的,对应为甲类
                    '吕说:中草药的需传1
                    If InStr(1, "7中草药", Nvl(!收费类别)) And Nvl(rsTemp!项目编码, "9000900099") = "9000900099" Then
                        StrInput = StrInput & "1" & "|"
                    Else
                        StrInput = StrInput & Nvl(rsTmp!收费等级) & "|"
                    End If
                    
                    If IsNull(rsTmp!收费类别) Then
                        StrInput = StrInput & Split(Get费用类别(Nvl(!收费类别)), "-")(0) & "|"
                    Else
                        StrInput = StrInput & Nvl(rsTmp!收费类别) & "|"
                    End If
                End If
                
                StrInput = StrInput & Format(rs明细("单价"), "0.0000") & "|"
                StrInput = StrInput & Format(rs明细("数量"), "0.00") & "|"
                StrInput = StrInput & Format(rs明细("实收金额"), "#####0.0000") & "|"         '金额
                
                StrInput = StrInput & ToVarchar(rsTemp("计算单位"), 20) & "|"      '单位
                StrInput = StrInput & ToVarchar(rsTemp("规格"), 14) & "|"
                StrInput = StrInput & ToVarchar(rsTemp("剂型"), 20) & "|"
                StrInput = StrInput & strDate & "|"
                
                 '0 表示开始循环，2 表示结束循环：在结束循环后才提交
                If rs明细.AbsolutePosition = 1 Then
                    If rs明细.AbsolutePosition = rs明细.RecordCount Then
                        StrInput = StrInput & "1"
                    Else
                        StrInput = StrInput & 0
                    End If
                ElseIf rs明细.AbsolutePosition = rs明细.RecordCount Then
                    StrInput = StrInput & 2
                Else
                    StrInput = StrInput & "1"
                End If
                
                If 业务请求_黔南(录入处方明细, StrInput, strOutput) = False Then
                    Exit Function
                End If
                If Not bln虚拟 Then
                    '不是虚拟结算，需确定相关的上传标志值
                    '审批编号
                    '为病人费用记录打上标记，以便随时上传
                    'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                    '摘要值:交易流水号|审批编号|住院(门诊)号|处方号|实际交易单价|实际等级
                    strArr = Split(strOutput, "|")  '--实际单价|实际等级|交易流水号
                    
                    strOutput = strArr(3) & "|" & str审批编号 & "|" & g病人身份_黔南.门诊号 & "|" & g病人身份_黔南.门诊号 & "|" & Val(strArr(1)) & "|" & strArr(2)
                    gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strOutput & "')"
                    zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
                End If
            End If
            g病人身份_黔南.费用总额 = g病人身份_黔南.费用总额 + Nvl(rs明细!实收金额, 0)
            rs明细.MoveNext
        Loop
    End With
    门诊明细写入 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Function 检查中心维护单价(ByVal lng病人ID As Long, ByVal lng细目ID As Long, ByVal dbl单价 As Double) As Boolean
    '检查中心维护中的单价是否与HIS输入的一致
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    检查中心维护单价 = False
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "" & _
        "   Select * From 医保收费目录 " & _
        "   Where (类别,编码) in (select 附注,项目编码 From 保险支付项目 where 险类=[1] and 收费细目id=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查中心维护单价", TYPE_黔南, lng细目ID)
    If rsTemp.EOF Then
        检查中心维护单价 = True
        'MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Nvl(rsTemp!维护标志) <> STR_中心维护代码 Then
        '刘兴宏:董杰要求更改.
        '确定明细项目单价
        '   医保编号|费用类别|医院单价
        StrInput = Nvl(rsTemp!编码) & "|"
        StrInput = StrInput & Nvl(rsTemp!收费类别) & "|"
        StrInput = StrInput & Format(dbl单价, "0.0000")
        
        If 业务请求_黔南(确定明细项目单价, StrInput, strOutput) = False Then
            Exit Function
        End If
        strOutput = Split(strOutput & "|", "|")(1)
        If Val(strOutput) <> dbl单价 Then
            gstrSQL = "Select 名称 From 收费细目 where ID=" & lng细目ID
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取细目名称"
            ShowMsgbox "存在项目“" & rsTemp!名称 & "”的单价不一样：" & vbCrLf & " 　医院:" & Format(dbl单价, "0.0000") & vbCrLf & "　中心:" & Format(Val(strOutput), "0.0000")
            Exit Function
        End If
    End If
    检查中心维护单价 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 门诊虚拟结算_黔南(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Static str上次门诊号 As String
    Dim str医保号 As String, StrInput As String, strOutput  As String
    Dim dbl个人帐户 As Double, strMessage As String
    Dim lng病人ID As Long, str规格 As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    Dim strArr
    Dim strDate As String
    
    On Error GoTo errHandle
    
    If rs明细.RecordCount = 0 Then
        str结算方式 = "个人帐户;0;0"
        门诊虚拟结算_黔南 = True
        Exit Function
    End If
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID")
    
    If g病人身份_黔南.病人ID <> lng病人ID Then
        MsgBox "该病人还没有经过身份验证，不能进行医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    '首先退掉以前发生的所有未结的费用，包括多次执行预结算
    If str上次门诊号 = g病人身份_黔南.门诊号 Then
        '已经赋值，说明该病人进行过预算
        StrInput = str上次门诊号 & "|" & str上次门诊号
        If 业务请求_黔南(处方退方, StrInput, strOutput) = False Then
            'Exit Function
        End If
    End If
    
    Dim rsTmp As New ADODB.Recordset
    Dim dbl单价 As Long
    
    '需先检查单价符合情况
    With rs明细
        Do While Not .EOF
            If 检查中心维护单价(Nvl(!病人ID, 0), Nvl(!收费细目ID, 0), Nvl(!单价, 0)) = False Then Exit Function
            .MoveNext
        Loop
    End With
'    str上次门诊号 = g病人身份_黔南.门诊号
    '然后插入处方明细
    If 门诊明细写入(rs明细, True) = False Then Exit Function
    
    '保存该值
    str上次门诊号 = g病人身份_黔南.门诊号
    '调用预结算
    '    交易特定输入数据：   住院（门诊）号|帐户消费标志
    '    交易特定输出数据:   费用总额|统筹支付|账户支付|现金支付|大病垫付
    '3． 账户消费标志 0 不从账户消费（即标志为0，不下个人帐户）， 1  使用系统参数设置值（即标志为1，下个人帐户）。
    StrInput = g病人身份_黔南.门诊号
    StrInput = StrInput & "|" & IIf(g病人身份_黔南.下个人帐户, "1", "0")
    
    If 更新就诊信息_黔南(0, strOutput) = False Then Exit Function
    If 业务请求_黔南(预结算, StrInput, strOutput) = False Then
        '出错,需对原处方进行退方
        '已经赋值，说明该病人进行过预算
        StrInput = g病人身份_黔南.门诊号 & "|" & g病人身份_黔南.门诊号
        If 业务请求_黔南(处方退方, StrInput, strOutput) = False Then
            Exit Function
        End If
    End If
    strArr = Split(strOutput, "|")
    
    str结算方式 = "个人帐户;" & Val(strArr(3)) & ";0"  '不能修改个人帐户，因为结算时已经不再传金额到前置机了
    
    If Val(strArr(2)) > 0 Then
        str结算方式 = str结算方式 & "|医保基金;" & Val(strArr(2)) & ";0"
    End If
    If Val(strArr(5)) > 0 Then
        str结算方式 = str结算方式 & "|大病垫付;" & Val(strArr(5)) & ";0"
    End If
        
    门诊虚拟结算_黔南 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function Get扩展信息(ByVal lng住院状态 As String, Optional str出院日期 As String) As String
    '获取扩展信息
    Dim strTemp As String
    Dim strArr
    Dim i As Integer
    If g病人身份_黔南.扩展信息 = "" Then Exit Function
    
    strArr = Split(g病人身份_黔南.扩展信息, "|")
    '定点医院1|定点医院2|定点医院3|定点医院4|定点医院5|出院日期|住院状态（1—住院；2—出院）|就诊医院|医疗类别
    If str出院日期 = "" Then
    Else
        strArr(5) = str出院日期
    End If
    strArr(6) = lng住院状态
    'strArr(7) = InitInfor_黔南.医院编码
    'strArr(8) = g病人身份_黔南.医疗类别
    
    For i = 0 To UBound(strArr)
        strTemp = strTemp & "|" & strArr(i)
    Next
    Get扩展信息 = Mid(strTemp, 2)
End Function


Public Function 门诊结算_黔南(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
        '此时所有收费细目必然有对应的医保编码
    Dim StrInput As String, strOutput As String
    Dim lng病人ID  As Long
    Dim dbl费用总额 As Double
    Dim str操作员 As String, strArr
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim str动态信息 As String
    Dim str交易流水号 As String
    Static str结算时间 As String
    Static oldlng病人ID As Long
    Dim lng就诊次数 As Long
        
    
    Dim datCurr As Date
    On Error GoTo errHandle
    Call DebugTool("进入门诊结算")
    gstrSQL = "Select a.*,a.付数*a.数次 as 数量,a.实收金额/(nvl(a.付数,1)*nvl(a.数次,1)) as 单价 From 门诊费用记录 a Where nvl(a.实收金额,0)<>0 and  结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取明细记录", lng结帐ID)
    If rs明细.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "没有填写收费记录"
        Exit Function
    End If

    lng病人ID = rs明细("病人ID")
    str操作员 = ToVarchar(IIf(IsNull(rs明细("操作员姓名")), UserInfo.姓名, rs明细("操作员姓名")), 20)
    
    If g病人身份_黔南.病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人还没有经过身份验证，不能进行医保结算。"
        Exit Function
    End If
        
    If lng病人ID = oldlng病人ID And str结算时间 = Format(rs明细!登记时间, "yyyy-mm-dd HH:MM:SS") Then
            '需要重新分配一个新号给门诊
            gstrSQL = "Select nvl(就诊次数,0)+1 as 就诊次数 From 保险帐户 where 险类=" & TYPE_黔南 & " and 病人id=" & lng病人ID
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取就诊次数"
            
            '更新保险帐户
            lng就诊次数 = Nvl(rsTemp!就诊次数, 1)
            g病人身份_黔南.门诊号 = lng病人ID & "_" & lng就诊次数
            
            '更新保险帐户
            gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_黔南 & ",'就诊次数','" & lng就诊次数 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊次数")
    
            '进行门诊登记
            '交易特定输入数据：住院（门诊）号|医保证号码|IC卡号|入院日期|入院疾病名称|经办人
            With g病人身份_黔南
                StrInput = .门诊号 & "|"
                StrInput = StrInput & .医保证号 & "|"
                StrInput = StrInput & .卡号 & "|"
                StrInput = StrInput & "" & "|"
                StrInput = StrInput & "" & "|"
                StrInput = StrInput & gstrUserName
            End With
            If 业务请求_黔南(病人登记, StrInput, strOutput) = False Then Exit Function
            If 更新就诊信息_黔南(0, strOutput) = False Then Exit Function
    Else
        '由于虚拟结算时，已经传了一次,但由于无法保存费用明细中的交易流水号,所以需先作废掉，再上传
        StrInput = g病人身份_黔南.门诊号 & "|" & g病人身份_黔南.门诊号
        If 业务请求_黔南(处方退方, StrInput, strOutput) = False Then
            Exit Function
        End If
    End If
    str结算时间 = Format(rs明细!登记时间, "yyyy-mm-dd HH:MM:SS")
    oldlng病人ID = Nvl(rs明细!病人ID, 0)

    '写入明细
    If 门诊明细写入(rs明细, False) = False Then Exit Function
    
    
    
    '调用结算
    Call DebugTool("准备调用门诊结算")
    '交易特定输入数据：  交易类型|住院(门诊)号|单据号|操作员姓名|帐户消费标志
    '交易特定输出数据:  费用总额|统筹支付|账户支付|现金支付|大病垫付|所有16项动态信息|交易流水号
    '交易类型定义如下：
    '   1正常结算 (出院结算)
    '   0住院中途结算
    '   -1反结算
    '   -2IC卡挂失后出院结算，本次结算（只针对住院）将所有费用转为现金支付，待到医保中心报销。
    '   账户消费标志 0 不从账户消费（即标志为0，不下个人帐户）， 1  使用系统参数设置值（即标志为1，下个人帐户）
    
    StrInput = "1|"
    StrInput = StrInput & g病人身份_黔南.门诊号 & "|"
    StrInput = StrInput & g病人身份_黔南.门诊号 & "|"
    StrInput = StrInput & str操作员 & "|"
    StrInput = StrInput & IIf(g病人身份_黔南.下个人帐户, "1", "0")
    If 业务请求_黔南(结算, StrInput, strOutput) = False Then
        Exit Function
    End If
    Call DebugTool("保险结算记录")
    
    '保存结算记录
    '---------------------------------------------------------------------------------------------
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
            
    Dim cur统筹支付 As Double
    Dim cur公务员补助 As Double
    Dim cur大病垫付 As Double
    strArr = Split(strOutput, "|")
    
    dbl费用总额 = Round(g病人身份_黔南.费用总额, 2)
    
    cur统筹支付 = Val(strArr(2))
    cur个人帐户 = Val(strArr(3))
    cur大病垫付 = Val(strArr(5))
    Dim i As Integer
    str动态信息 = ""
    
    '获取动态信息
    For i = 6 To UBound(strArr) - 1
        str动态信息 = str动态信息 & "|" & strArr(i)
    Next
    str动态信息 = Mid(str动态信息, 2)

    
    '进行写卡
    '设置动态属性
    If sCard_属性值(2, str动态信息, True) = False Then
        GoTo Err反结算:
    End If
    
    '写扩展信息
    '定点医院1|定点医院2|定点医院3|定点医院4|定点医院5|出院日期|住院状态（1—住院；2—出院）|就诊医院|医疗类别
    'bytType:1-SiCardBaseInfo社保卡基本信息
    '        2-SiCardDynaInfo社保卡动态信息
    '        3-SiCardAcctInfo社保卡帐户信息
    '        4-SiCardExtInfo社保卡扩展信息
    
    
    If sCard_属性值(4, Get扩展信息("2"), True) = False Then
        GoTo Err反结算:
    End If
    str交易流水号 = strArr(UBound(strArr))
    
    If sCard_SaveCard = False Then GoTo Err反结算:
     
    
    '帐户年度信息
    datCurr = zlDatabase.Currentdate
    
    Call Get帐户信息(TYPE_黔南, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
                
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_黔南 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 + cur统筹支付 & "," & _
        cur统筹报销累计 + cur统筹支付 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存帐户年度信息")
    
    
    Dim dbl帐户余额 As Double
    If Save保险帐户_帐户余额(lng病人ID, str动态信息, dbl帐户余额) = False Then
        GoTo Err反结算:
    End If
        
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(帐户增加累计),帐户累计支出_IN(帐户累计支出),累计进入统筹_IN(累计进入统筹_IN),累计统筹报销_IN(累计统筹报销),住院次数_IN(住院次数累计),起付线(无),封顶线_IN(无),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(无),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算时产生流水号),主页ID_IN,中途结帐_IN,备注_IN
    
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    '|费用总额|统筹支付|账户支付|现金支付|大病垫付|所有16项动态信息|交易流水号
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_黔南 & "," & lng病人ID & "," & Year(datCurr) & "," & _
            cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & cur统筹报销累计 & "," & int住院次数累计 & ",0,0," & IIf(g病人身份_黔南.下个人帐户, 1, 0) & "," & _
            dbl费用总额 & "," & dbl帐户余额 & " ,0," & _
            cur统筹支付 & "," & cur统筹支付 & ",0,0," & cur个人帐户 & ",'" & _
            str交易流水号 & "',NULL,NULL,'" & g病人身份_黔南.门诊号 & "|" & str动态信息 & "')"
            
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    '---------------------------------------------------------------------------------------------
    门诊结算_黔南 = True
    Exit Function
Err反结算:
    '设置失败:进行反结算
    '进行反结算
    '交易特定输入数据：  交易类型|住院(门诊)号|单据号|操作员姓名
    '交易特定输出数据:  费用总额|统筹支付|账户支付|现金x支付|大病垫付|所有16项动态信息|交易流水号
    '交易类型定义如下：
    '   1正常结算 (出院结算)
    '   0住院中途结算
    '   -1反结算
    '   -2IC卡挂失后出院结算，本次结算（只针对住院）将所有费用转为现金支付，待到医保中心报销。

    StrInput = "-1|"
    StrInput = StrInput & g病人身份_黔南.门诊号 & "|"
    StrInput = StrInput & g病人身份_黔南.门诊号 & "|"
    StrInput = StrInput & str操作员 & "|"
    StrInput = StrInput & IIf(g病人身份_黔南.下个人帐户, "1", "0")
    Call 业务请求_黔南(结算, StrInput, strOutput)
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function 更新就诊信息_黔南(ByVal bytType As Byte, strOutPutstring As String, Optional ByVal str入院日期 As String = "", _
        Optional ByVal str出院日期 As String = "", _
        Optional bln反结算 As Boolean = False, Optional blnNOWriteCard As Boolean = False) As Boolean
        'bytType:0-门诊,1-入院登记,2-住院结算
        Dim StrInput As String, i As Integer, strTemp As String
        Dim strArr
        Dim strDate As String
        '更新就诊信息
        更新就诊信息_黔南 = False
        Err = 0
        On Error GoTo errHand:
        Select Case bytType
            Case 0 '门诊
                '只更新医疗类别,
                '住院号|更新标志|医疗类别|入院日期|入院疾病名称|出院日期|确诊疾病编码|经办人|所有16位动态信息
                '返回:更新动态信息标志|所有16位动态信息
                strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
                StrInput = g病人身份_黔南.门诊号 & "|"
                If bln反结算 Then
                    '只更新动态信息和操作员
                    StrInput = StrInput & "0000011111111111111111" & "|"
                ElseIf g病人身份_黔南.病种编码 = "" Or g病人身份_黔南.病种编码 = "000000" Then
                    '不更新诊断
                    StrInput = StrInput & "1101011111111111111111" & "|"
                    StrInput = StrInput & g病人身份_黔南.医疗类别 & "|"
                    StrInput = StrInput & strDate & "|"
                    StrInput = StrInput & strDate & "|"
                Else
                    StrInput = StrInput & "1111111111111111111111" & "|"
                    StrInput = StrInput & g病人身份_黔南.医疗类别 & "|"
                    StrInput = StrInput & strDate & "|"
                    StrInput = StrInput & g病人身份_黔南.病种名称 & "|"
                    StrInput = StrInput & strDate & "|"
                    StrInput = StrInput & g病人身份_黔南.病种编码 & "|"
                End If
                StrInput = StrInput & gstrUserName & "|"
                StrInput = StrInput & g病人身份_黔南.动态信息
        Case 1  '入院登记
                '住院号|更新标志|医疗类别|入院日期|入院疾病名称|出院日期|确诊疾病编码|经办人|所有16位动态信息
                '返回:更新动态信息标志|所有16位动态信息
                
                '3． 确诊疾病编码 : 这里指中心编码，门诊可以不置，但是住院必须提供有效的编码。
                '4. 入院登记时调用'更新就诊信息'时，不更新'出院日期'；出院结算时，不更新'入院日期''疾病名称'
                
                strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
                StrInput = g病人身份_黔南.住院号 & "|"
                
                If g病人身份_黔南.病种编码 = "" Or g病人身份_黔南.病种编码 = "000000" Then
                    '不更新诊断
                    StrInput = StrInput & "1100011111111111111111" & "|"
                    StrInput = StrInput & g病人身份_黔南.医疗类别 & "|"
                    StrInput = StrInput & str入院日期 & "|"
                Else
                    StrInput = StrInput & "1110111111111111111111" & "|"
                    StrInput = StrInput & g病人身份_黔南.医疗类别 & "|"
                    StrInput = StrInput & strDate & "|"
                    StrInput = StrInput & g病人身份_黔南.病种名称 & "|"
                    StrInput = StrInput & g病人身份_黔南.病种编码 & "|"
                End If
                StrInput = StrInput & gstrUserName & "|"
                StrInput = StrInput & g病人身份_黔南.动态信息
        Case 2      '住院结算
    
            '更新就诊信息(目前只改出院日期,确诊编码,经办人)
            '住院号|更新标志|医疗类别|入院日期|入院疾病名称|出院日期|确诊疾病编码|经办人|所有16位动态信息
            StrInput = g病人身份_黔南.住院号 & "|"
            
            If bln反结算 Then
                    '只更新动态信息和操作员
                    StrInput = StrInput & "0000011111111111111111" & "|"
            ElseIf g病人身份_黔南.病种编码 = "" Or g病人身份_黔南.病种编码 = "000000" Then
                StrInput = StrInput & "0001011111111111111111" & "|"
                StrInput = StrInput & str出院日期 & "|"
            Else
                StrInput = StrInput & "0001111111111111111111" & "|"
                StrInput = StrInput & str出院日期 & "|"
                StrInput = StrInput & g病人身份_黔南.病种编码 & "|"
            End If
            StrInput = StrInput & gstrUserName & "|"
            StrInput = StrInput & g病人身份_黔南.动态信息
        End Select
        
        If 业务请求_黔南(更新就诊信息, StrInput, strOutPutstring) = False Then Exit Function

        '7． 更新动态信息标志：如果该标志提示需要更新卡动态信息，则需要马上使用读写卡组件对卡上的动态信息进行更新。
        '   1 需要更新卡动态信息，使用紧接其后的动态信息值更新卡上的动态信息
        '   0 不需要更新卡动态信息
        strArr = Split(strOutPutstring, "|")
        If Val(strArr(1)) = 1 And blnNOWriteCard = False Then
            '需立即更新动态信息
            'bytType:1-SiCardBaseInfo社保卡基本信息
            '        2-SiCardDynaInfo社保卡动态信息
            '        3-SiCardAcctInfo社保卡帐户信息
            '        4-SiCardExtInfo社保卡扩展信息
            strTemp = ""
            For i = 2 To UBound(strArr)
                strTemp = strTemp & "|" & strArr(i)
            Next
            strTemp = Mid(strTemp, 2)
            sCard_属性值 2, strTemp, True
            sCard_SaveCard
        End If
        更新就诊信息_黔南 = True
        Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function



Public Function 门诊结算冲销_黔南(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput  As String, str流水号 As String
    Dim lng冲销ID As Long
    Dim strArr
    Dim rs明细 As New ADODB.Recordset
    Dim i As Long
    
    Dim dbl费用总额 As Double, dbl个人帐户 As Double
    Dim dbl帐户增加累计 As Currency, dbl帐户支出累计 As Currency
    Dim dbl进入统筹累计 As Currency, dbl统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim str动态信息 As String
    Dim curDate As Date
    Dim str医保证号 As String
    
    On Error GoTo errHandle
    
    curDate = zlDatabase.Currentdate
    If Get病人信息(lng病人ID) = False Then Exit Function
    str医保证号 = g病人身份_黔南.医保证号
    If 获取参保人员信息_黔南() = False Then Exit Function
    If str医保证号 <> g病人身份_黔南.医保证号 Then
        Err.Raise 9000, gstrSysName, "卡插入错误!"
        Exit Function
    End If
    
    
    gstrSQL = "select 备注 from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取原来的结算记录", TYPE_黔南, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    g病人身份_黔南.门诊号 = Split(Nvl(rsTemp!备注) & "|", "|")(0)
    
    
    If 更新就诊信息_黔南(0, strOutput, , , True) = False Then
        Err.Raise 9000, gstrSysName, "更新就诊信息失败!"
        Exit Function
    End If
    
    gstrSQL = "Select * From 门诊费用记录  " & _
        " Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取冲销记录", lng结帐ID)
    
    Do Until rs明细.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        dbl费用总额 = dbl费用总额 + Nvl(rs明细("结帐金额"), 0)
        rs明细.MoveNext
    Loop
    dbl费用总额 = Round(dbl费用总额, 2)
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")
    
    
    
    gstrSQL = "Select * From 门诊费用记录 " & _
        " Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取冲销记录", lng冲销ID)
    Do While Not rsTemp.EOF
        rs明细.Filter = 0
        rs明细.Filter = "NO='" & Nvl(rsTemp!NO) & "' and 记录性质=" & Nvl(rsTemp!记录性质) & " and 序号=" & Nvl(rsTemp!序号)
        If rs明细.EOF Then
            ShowMsgbox "冲销有一条以上的冲销明细未找到!"
            Exit Function
        End If
        
        '更新上传标志
        gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(rsTemp!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(rs明细!摘要) & "')"
        zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取原来的结算记录", TYPE_黔南, lng结帐ID)
    
    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    
    str流水号 = rsTemp("支付顺序号")
    
    '进行反结算
    '交易特定输入数据：  交易类型|住院(门诊)号|单据号|操作员姓名
    '交易特定输出数据:  费用总额|统筹支付|账户支付|现金x支付|大病垫付|所有16项动态信息|交易流水号
    '交易类型定义如下：
    '   1正常结算 (出院结算)
    '   0住院中途结算
    '   -1反结算
    '   -2IC卡挂失后出院结算，本次结算（只针对住院）将所有费用转为现金支付，待到医保中心报销。
    g病人身份_黔南.门诊号 = Split(Nvl(rsTemp!备注) & "|", "|")(0)
    StrInput = "-1|"
    StrInput = StrInput & g病人身份_黔南.门诊号 & "|"
    StrInput = StrInput & g病人身份_黔南.门诊号 & "|"
    StrInput = StrInput & gstrUserName & "|" & IIf(Nvl(rsTemp!实际起付线, 0) = 1, 1, 0)
    
    If 业务请求_黔南(结算, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    strArr = Split(strOutput, "|")
    Dim dbl统筹支付 As Double
    Dim dbl大病垫付 As Double
    
    dbl统筹支付 = Val(strArr(2))
    dbl个人帐户 = Val(strArr(3))
    dbl大病垫付 = Val(strArr(5))
    If Abs(dbl个人帐户) <> Abs(Nvl(rsTemp!个人帐户支付, 0)) Then
        ShowMsgbox "本次中心冲帐的个人帐户支付不等于上次结算的个人帐户支付!"
        Exit Function
    End If
    
    str动态信息 = ""
    '获取动态信息
    For i = 6 To UBound(strArr) - 1
        str动态信息 = str动态信息 & "|" & strArr(i)
    Next
    
    str动态信息 = Mid(str动态信息, 2)
    
   
         
         
    '进行写卡
    '设置动态属性
    If sCard_属性值(2, str动态信息, True) = False Then
        GoTo Err冲正:
    End If
    
    '写扩展信息
    '定点医院1|定点医院2|定点医院3|定点医院4|定点医院5|出院日期|住院状态（1—住院；2—出院）|就诊医院|医疗类别
    'bytType:1-SiCardBaseInfo社保卡基本信息
    '        2-SiCardDynaInfo社保卡动态信息
    '        3-SiCardAcctInfo社保卡帐户信息
    '        4-SiCardExtInfo社保卡扩展信息
    
    '读卡
    
    If sCard_属性值(4, Get扩展信息("2"), True) = False Then
        GoTo Err冲正:
    End If
    str流水号 = strArr(UBound(strArr))
    If sCard_SaveCard = False Then GoTo Err冲正:
    
    
    
    '退处方单
    StrInput = g病人身份_黔南.门诊号 & "|" & g病人身份_黔南.门诊号
    If 业务请求_黔南(处方退方, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_黔南, lng病人ID, Year(curDate), int住院次数累计, dbl帐户增加累计, dbl帐户支出累计, dbl进入统筹累计, dbl统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_黔南 & "," & Year(curDate) & "," & _
        dbl帐户增加累计 & "," & dbl帐户支出累计 - dbl个人帐户 & "," & dbl进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        dbl统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "插入帐户年度信息")
    
    Dim dbl帐户余额 As Double
    If Save保险帐户_帐户余额(lng病人ID, str动态信息, dbl帐户余额) = False Then
        GoTo Err冲正:
    End If
    
    '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(帐户增加累计),帐户累计支出_IN(帐户累计支出),累计进入统筹_IN(累计进入统筹_IN),累计统筹报销_IN(累计统筹报销),住院次数_IN(住院次数累计),起付线(无),封顶线_IN(无),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(无),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算时产生流水号),主页ID_IN,中途结帐_IN,备注_IN
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_黔南 & "," & lng病人ID & "," & Year(curDate) & "," & _
        dbl帐户增加累计 & "," & dbl帐户支出累计 - dbl个人帐户 & "," & dbl进入统筹累计 & "," & dbl统筹报销累计 & "," & int住院次数累计 & ",0,0," & Nvl(rsTemp!实际起付线, 0) & "," & _
        dbl费用总额 * -1 & "," & dbl帐户余额 & ",0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0,0," & dbl个人帐户 * -1 & ",'" & _
       str流水号 & "',NULL,0,'" & g病人身份_黔南.门诊号 & "|" & str动态信息 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新保险结算信息")
    门诊结算冲销_黔南 = True
    Exit Function
Err冲正:
    Call 业务请求_黔南(冲正交易, str流水号 & "|10|" & gstrUserName, strOutput)
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function 病人入院登记处理(lng病人ID As Long, lng主页ID As Long) As Boolean
    '进行门诊登记
    Dim StrInput As String, strOutput As String
    Dim str交易流水号 As String
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    '交易特定输入数据：住院（门诊）号|医保证号码|IC卡号|入院日期|入院疾病名称|经办人
    '交易特定输出数据:       执行成功时返回交易流水号 , 执行失败时为失败原因描述
    gstrSQL = "Select C.住院号,C.当前床号,to_char(A.确诊日期,'yyyy-MM-dd') as 确诊日期,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') 入院经办时间," & _
        " to_char(A.登记时间,'yyyy-mm-dd') 入院日期  ,to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') 入院时间,D.入院诊断编码,D.入院诊断名称,G.确诊诊断编码,g.确诊诊断名称 " & _
        " From 病案主页 A,部门表 B,病人信息 C, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,b.编码,'')) AS 入院诊断编码,max(DECODE(a.诊断次序,1,b.名称,'')) AS 入院诊断名称 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =1 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by  病人id,主页id)   D," & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,2,b.编码,'')) AS 确诊诊断编码,max(DECODE(a.诊断次序,2,b.名称,'')) AS 确诊诊断名称 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =1 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by  病人id,主页id)   g" & _
        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
        "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
        "       and A.主页id=g.主页id(+) and a.病人id=g.病人id(+) " & _
        ""
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取入院信息"
        
    With g病人身份_黔南
        .入院诊断编码 = Nvl(rsTemp!入院诊断编码)
        .入院诊断名称 = Nvl(rsTemp!入院诊断名称)
        .确诊诊断编码 = Nvl(rsTemp!确诊诊断编码)
        .确诊诊断名称 = Nvl(rsTemp!确诊诊断名称)
        
        StrInput = .住院号 & "|"
        StrInput = StrInput & .医保证号 & "|"
        StrInput = StrInput & .卡号 & "|"
        StrInput = StrInput & Nvl(rsTemp!入院日期) & "|"
        StrInput = StrInput & .病种名称 & "|"
        'strInput = strInput & Nvl(rsTemp!入院诊断名称) & "|"
        StrInput = StrInput & Nvl(rsTemp!经办人, gstrUserName)
    End With
    
    Err = 0
    On Error GoTo errHand:
    
    If 业务请求_黔南(病人登记, StrInput, strOutput) = False Then Exit Function
    str交易流水号 = Split(strOutput, "|")(2)
      
    '保存将交易流水号
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_黔南 & ",'顺序号','''" & str交易流水号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存交易流水号")
    
    
    If 更新就诊信息_黔南(1, strOutput, Nvl(rsTemp!入院日期)) = False Then
        GoTo Err冲正:
    End If
  
     
    '写扩展信息
    '定点医院1|定点医院2|定点医院3|定点医院4|定点医院5|出院日期|住院状态（1—住院；2—出院）|就诊医院|医疗类别
    'bytType:1-SiCardBaseInfo社保卡基本信息
    '        2-SiCardDynaInfo社保卡动态信息
    '        3-SiCardAcctInfo社保卡帐户信息
    '        4-SiCardExtInfo社保卡扩展信息
  '  str出院日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If sCard_属性值(4, Get扩展信息("1", ""), True) = False Then
        GoTo Err冲正:
        Exit Function
    End If
    
    '需改变住院状态
    If sCard_SaveCard = False Then
        GoTo Err冲正:
        Exit Function
    End If
    
    
    病人入院登记处理 = True
    Exit Function
Err冲正:
        '被冲正交易流水号|被冲正交易类型代码|操作员代码
        StrInput = str交易流水号 & "|"
        StrInput = StrInput & "01" & "|"
        StrInput = StrInput & gstrUserName
        If 业务请求_黔南(冲正交易, StrInput, strOutput) = False Then
        End If
        Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记_黔南(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutput As String, StrInput As String
    Dim lng就诊次数 As Long
    Dim str交易流水号 As String
    '获取住院号
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select nvl(就诊次数,0)+1 as 就诊次数 From 保险帐户 where 病人ID=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "入院登记"
    lng就诊次数 = Nvl(rsTemp!就诊次数, 1)
    g病人身份_黔南.住院号 = lng病人ID & "_" & lng就诊次数
    
    '更新保险帐户
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_黔南 & ",'就诊次数','" & lng就诊次数 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊次数")
    
    '先进行登记处理
    If 病人入院登记处理(lng病人ID, lng主页ID) = False Then
        Exit Function
    End If
    
    
    '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_黔南 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_黔南 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_黔南 = False
End Function
Private Function Get交易代码(ByVal intType As 业务类型_黔南, Optional bln读名称 As Boolean = False) As String
    Select Case intType
        Case 病人登记
            Get交易代码 = IIf(bln读名称, "病人登记", "01")
        Case 项目审批结果查询
            Get交易代码 = IIf(bln读名称, "项目审批结果查询", "02")
        Case 获得通过的审批信息
            Get交易代码 = IIf(bln读名称, "获得通过的审批信息", "03")
        Case 参保资格审查
            Get交易代码 = IIf(bln读名称, "参保资格审查", "04")
        Case 更新就诊信息
            Get交易代码 = IIf(bln读名称, "更新就诊信息", "05")
        Case 录入处方明细
            Get交易代码 = IIf(bln读名称, "录入处方明细", "06")
        Case 确定明细项目单价
            Get交易代码 = IIf(bln读名称, "确定明细项目单价", "07")
        Case 处方退方
            Get交易代码 = IIf(bln读名称, "处方退方", "08")
        Case 预结算
            Get交易代码 = IIf(bln读名称, "预结算", "09")
        Case 结算
            Get交易代码 = IIf(bln读名称, "结算", "10")
        Case 获得参保人员信息
            Get交易代码 = IIf(bln读名称, "获得参保人员信息", "13")
        Case 冲正交易
            Get交易代码 = IIf(bln读名称, "冲正交易", "99")
        Case 收费目录下载处理
            Get交易代码 = IIf(bln读名称, "收费目录下载处理", "02")
        Case 收费目录下载预处理
            Get交易代码 = IIf(bln读名称, "收费目录下载预处理", "01")
        Case 读取卡内数据
            Get交易代码 = IIf(bln读名称, "读取卡内数据", "-1")
        Case 读卡组件_扩展信息
            Get交易代码 = IIf(bln读名称, "读卡组件_扩展信息", "-1")
        Case Else
            Get交易代码 = IIf(bln读名称, "错误的交易代码", "-1")
    End Select
End Function
Public Function 业务请求_黔南(ByVal intType As 业务类型_黔南, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:strOutPutString-输出串,按参数顺序,以tab键分隔的返回串
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, lngReturn As Long, strOutput As String, strReturn As String
    Dim str交易代码 As String
    Dim str交易名称 As String
    Dim i As Integer
    Dim strArr
    
    str交易代码 = Get交易代码(intType)
    StrInput = str交易代码 & "|" & strInputString
    str交易名称 = str交易代码 & "—" & Get交易代码(intType, True)
    
    DebugTool "进入业务请求函数(业务类型为:" & str交易名称 & ") " & vbCrLf & ".....输入参数为" & Trim(StrInput)
    
    业务请求_黔南 = False
    If InitInfor_黔南.模拟数据 Then
        '读取模拟数据
        Read模拟数据 intType, StrInput, strOutPutstring
         业务请求_黔南 = True
        Exit Function
    End If
    strOutput = Space(5000)
    
    Err = 0: On Error GoTo errHand:
'    If gobjTest Is Nothing Then
'        Set gobjTest = CreateObject("SiInterface.clsSiInterface")
'    End If
    Select Case intType
        Case 医保初始化
            lngReturn = Init()
            If lngReturn <> 0 Then
                MsgBox "不能正确调用初始化医保接口。", vbInformation, gstrSysName
                Exit Function
            End If
        Case 收费目录下载处理, 收费目录下载预处理
            lngReturn = QUERY_HANDLE(StrInput, strOutput)
            '4、 返回0表示执行成功，返回-1表示执行失败，返回100表示待查项目不存在。
            If lngReturn = -1 Then
                ShowMsgbox "下载失败!"
                Exit Function
            End If
            If lngReturn = 100 Then
                ShowMsgbox "待查项目不存在!"
                Exit Function
            End If
        Case Else
            '
            '获得参保人员信息, 病人登记, 参保资格审查
            lngReturn = BUSINESS_HANDLE(StrInput, strOutput)
            strOutput = Trim(TruncZero(strOutput))
            strArr = Split(strOutput, "|")
            '输出参数的前6位是业务执行代码。如果业务成功，执行代码为'     0'，下一个元素是交易流水号；如果业务失败，业务执行代码后的下一个元素是出错信息。
            If lngReturn <> 0 Then
                '业务调用失败
                strReturn = "医保接口出现警告：" & vbCrLf & strArr(0)
                ShowMsgbox strReturn
                Exit Function
            End If
    End Select
    strOutPutstring = "0|" & Trim(strOutput)
    业务请求_黔南 = True
    DebugTool ".....输出参数为(成功):" & Trim(strOutPutstring)
     Exit Function
errHand:
    DebugTool ".....输出参数为(失败):" & Trim(strOutPutstring)
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_黔南(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
            
    '刘兴宏:20040923增加的
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim str医保号  As String
    Dim str出院日期 As String
    
    Err = 0
    On Error GoTo errHand
    
    DebugTool "进入扩院登撤消接口"
    
    入院登记撤销_黔南 = False
    
    If 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "存在未结费用，不能撤消入院登记"
        Exit Function
    End If
    
    If 获取参保人员信息_黔南 = False Then Exit Function
    
    
    '写扩展信息
    '定点医院1|定点医院2|定点医院3|定点医院4|定点医院5|出院日期|住院状态（1—住院；2—出院）|就诊医院|医疗类别
    'bytType:1-SiCardBaseInfo社保卡基本信息
    '        2-SiCardDynaInfo社保卡动态信息
    '        3-SiCardAcctInfo社保卡帐户信息
    '        4-SiCardExtInfo社保卡扩展信息
    str出院日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If sCard_属性值(4, Get扩展信息("2", ""), True) = False Then
        Exit Function
    End If
    
    '需改变住院状态
    If sCard_SaveCard = False Then Exit Function
    
    
    
    '调用冲正交易
    gstrSQL = "Select 顺序号 From 保险帐户 where 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取入院登记时的交易流水号"
    If rsTemp.EOF Then
        ShowMsgbox "在医保中无此病人!"
        Exit Function
    End If
    '交易特定输入数据：  被冲正交易流水号|被冲正交易类型代码|操作员代码
    '交易特定输出数据:   未用
    
    StrInput = Nvl(rsTemp!顺序号) & "|"
    StrInput = StrInput & "01" & "|"
    StrInput = StrInput & gstrUserName
    If 业务请求_黔南(冲正交易, StrInput, strOutput) = False Then Exit Function
    '更新医保帐户
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_黔南 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    
    DebugTool "取消成功"
    入院登记撤销_黔南 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function 出院登记_黔南(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    '个人状态的修改
  
    '改变当前状态
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_黔南 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_黔南 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_黔南 = False
End Function
Public Function 出院登记撤销_黔南(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '出院登记撤消
     '改变病人状态
     If Not 存在未结费用(lng病人ID, lng主页ID) Then
            ShowMsgbox "该病人已经出院结算了,需对结算进行反结算!"
            Exit Function
     End If
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_黔南 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    出院登记撤销_黔南 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_黔南(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
  '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput As String

    Dim str操作员 As String
    Dim lng主页ID As Long
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim datCurr As Date
    Dim strArr
    Dim i As Integer
    Dim str动态信息  As String
    
    '医保病人必须先出院后结算:程池福加入20061120.
    If 医保病人已经出院(lng病人ID) = False Then
        Err.Raise 9000, gstrSysName, "医保病人必须先出院后才能办理结算!"
        Exit Function
    End If
        
    If g病人身份_黔南.病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人没有完成医保的预结算操作，不能进行结算。"
        Exit Function
    End If
        
    Err = 0: On Error GoTo errHand:
    Call DebugTool("进入住院结算")
    
    
    With g结算数据
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID)
        If IsNull(rsTemp("主页ID")) = True Then
            Err.Raise 9000, gstrSysName, "只有住院病人才可以使用医保结算。"
            Exit Function
        End If
        lng主页ID = rsTemp("主页ID")
    End With
    
    
    
    
    '再次预结，因为可能此期间又上传了明细数据
    '交易特定输入数据：   住院（门诊）号
    '交易特定输出数据:   费用总额|统筹支付|账户支付|现金支付|大病垫付
    
    Dim str结算方式  As String
    StrInput = g病人身份_黔南.住院号 & "|" & IIf(g病人身份_黔南.下个人帐户, "1", "0")
    
    If 业务请求_黔南(预结算, StrInput, strOutput) = False Then
        Exit Function
    End If
    strArr = Split(strOutput, "|")
    
    '检查虚拟结算是否一至
    With 虚拟结算数据
        If Round(.费用总额, 2) <> Round(Val(strArr(1)), 2) Or Round(.统筹支付, 2) <> Round(Val(strArr(2)), 2) Or _
            Round(.账户支付, 2) <> Round(Val(strArr(3)), 2) Or Round(.现金支付, 2) <> Round(Val(strArr(4)), 2) Or _
            Round(.大病垫付, 2) <> Round(Val(strArr(5)), 2) Then
            ShowMsgbox "本次结算时与虚拟结算不符,请检查..." & vbCrLf & _
                    "   费用总额:" & Format(.费用总额, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(Val(strArr(1)), "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    "   统筹支付:" & Format(.统筹支付, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(Val(strArr(2)), "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    "   账户支付:" & Format(.账户支付, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(Val(strArr(3)), "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    "   现金支付:" & Format(.现金支付, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(Val(strArr(4)), "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    "   大病垫付:" & Format(.大病垫付, "####0.00;####0.00;0.00;0.00") & vbTab & vbTab & Format(Val(strArr(5)), "####0.00;####0.00;0.00;0.00") & vbCrLf & _
                    ""
            Exit Function
        End If
    End With
    
    '正式结算
    '交易特定输入数据：  交易类型|住院(门诊)号|单据号|操作员姓名|帐户消费标志
    '交易特定输出数据:  费用总额|统筹支付|账户支付|现金支付|大病垫付|所有16项动态信息|交易流水号
    '交易类型定义如下：
    '   1正常结算 (出院结算)
    '   0住院中途结算
    '   -1反结算
    '   -2IC卡挂失后出院结算，本次结算（只针对住院）将所有费用转为现金支付，待到医保中心报销。
    '   账户消费标志 0 不从账户消费（即标志为0，不下个人帐户）， 1  使用系统参数设置值（即标志为1，下个人帐户）
    
    If g病人身份_黔南.中途结算 = True Then
        StrInput = 0 & "|"
    Else
        StrInput = 1 & "|"
    End If
    StrInput = StrInput & g病人身份_黔南.住院号 & "|"
    StrInput = StrInput & lng结帐ID & "|"
    StrInput = StrInput & gstrUserName & "|"
    StrInput = StrInput & IIf(g病人身份_黔南.下个人帐户, "1", "0")
    
    If 业务请求_黔南(结算, StrInput, strOutput) = False Then
        Exit Function
    End If
        
    strArr = Split(strOutput, "|")
    str动态信息 = ""
    '获取动态信息
    For i = 6 To UBound(strArr) - 1
        str动态信息 = str动态信息 & "|" & strArr(i)
    Next
    
    str动态信息 = Mid(str动态信息, 2)
    

    
    Dim objData As 结算数据
    With objData
        .费用总额 = Val(strArr(1))
        .统筹支付 = Val(strArr(2))
        .账户支付 = Val(strArr(3))
        .现金支付 = Val(strArr(4))
        .大病垫付 = Val(strArr(5))
        .动态信息 = str动态信息
        .交易流水号 = strArr(UBound(strArr))
    End With
    
    '写卡.
    '   如果病人账户余额不为零，根据结算函数返回的"账户支付"调用卡函数SaleTrans进行下帐操作。
    '   并且调用卡函数UpdateJrtclj, UpdateQfxzflj和卡函数UpdateDynInfo对卡动态信息进行更新操作。
    '   注意一定要调用卡函数UpdateHospStatus将病人的住院状态设置为2（出院状态），
    '   并调用卡函数AddHospTimes将病人的住院次数加一
    
    '??需加入写卡代码
    'bytType:1-SiCardBaseInfo社保卡基本信息
    '        2-SiCardDynaInfo社保卡动态信息
    '        3-SiCardAcctInfo社保卡帐户信息
    '        4-SiCardExtInfo社保卡扩展信息
    
    '进行写卡
    '设置动态属性
    If sCard_属性值(2, str动态信息, True) = False Then
        GoTo Err冲正:
    End If
    
    '写扩展信息
    '定点医院1|定点医院2|定点医院3|定点医院4|定点医院5|出院日期|住院状态（1—住院；2—出院）|就诊医院|医疗类别
    'bytType:1-SiCardBaseInfo社保卡基本信息
    '        2-SiCardDynaInfo社保卡动态信息
    '        3-SiCardAcctInfo社保卡帐户信息
    '        4-SiCardExtInfo社保卡扩展信息
    If g病人身份_黔南.中途结算 Then
    Else
        If sCard_属性值(4, Get扩展信息("2"), True) = False Then
            GoTo Err冲正:
        End If
    End If
    If sCard_SaveCard = False Then GoTo Err冲正:


    '填写结算表
    Call DebugTool("填写结算记录")
    datCurr = zlDatabase.Currentdate
    
    '帐户年度信息
    Call Get帐户信息(TYPE_黔南, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_黔南 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 & "," & _
        cur进入统筹累计 + objData.统筹支付 & "," & _
        cur统筹报销累计 + objData.统筹支付 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存帐户年度信息")
    
    
    Dim dbl帐户余额 As Double
    If Save保险帐户_帐户余额(lng病人ID, str动态信息, dbl帐户余额) = False Then
        GoTo Err冲正:
    End If
    
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(帐户增加累计),帐户累计支出_IN(帐户累计支出),累计进入统筹_IN(累计进入统筹_IN),累计统筹报销_IN(累计统筹报销),住院次数_IN(住院次数累计),起付线(无),封顶线_IN(无),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(当前帐户余额),首先自付金额_IN(无),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算时产生流水号),主页ID_IN,中途结帐_IN,备注_IN
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
   
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_黔南 & "," & lng病人ID & "," & Year(datCurr) & "," & _
            cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & cur统筹报销累计 & "," & int住院次数累计 & ",0,0," & IIf(g病人身份_黔南.下个人帐户, "1", "0") & "," & _
            g病人身份_黔南.费用总额 & "," & dbl帐户余额 & ",0," & _
            objData.统筹支付 & "," & objData.统筹支付 & ",0,0," & objData.账户支付 & ",'" & _
            objData.交易流水号 & "'," & lng主页ID & "," & IIf(g病人身份_黔南.中途结算, 1, 0) & ",'" & g病人身份_黔南.住院号 & "')"
            
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    '---------------------------------------------------------------------------------------------
    
    住院结算_黔南 = True
    Exit Function
Err冲正:
    Call 业务请求_黔南(冲正交易, objData.交易流水号 & "|10|" & gstrUserName, strOutput)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function


Public Function 住院结算冲销_黔南(lng结帐ID As Long) As Boolean
     '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput  As String
    Dim lng冲销ID As Long, str流水号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, i As Integer
    Dim curDate As Date
    Dim strArr
    Dim str动态信息  As String, str医保证号 As String
    Err = 0: On Error GoTo errHand:
    
    curDate = zlDatabase.Currentdate
    
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", TYPE_黔南, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    
   
    '判断病人的住院结算数据是否允许作废。判断标准是检查病人有新的住院记录，如果有，就不能交冲销
    If Can住院结算冲销(rsTemp("病人ID"), rsTemp("主页ID")) = False Then Exit Function
    
    If Get病人信息(rsTemp("病人id").Value) = False Then Exit Function
    
    str医保证号 = g病人身份_黔南.医保证号
    If 获取参保人员信息_黔南() = False Then Exit Function
    If str医保证号 <> g病人身份_黔南.医保证号 Then
        Err.Raise 9000, gstrSysName, "不是该医保病人的医保证"
        Exit Function
    End If
    
    str流水号 = rsTemp("支付顺序号")
    g病人身份_黔南.住院号 = Split(Nvl(rsTemp!备注, "|"), "|")(0)
     
     
    If 更新就诊信息_黔南(2, strOutput, , , True) = False Then
        Err.Raise 9000, gstrSysName, "更新就诊信息失败!"
        Exit Function
    End If
    
    g病人身份_黔南.下个人帐户 = IIf(Nvl(rsTemp!实际起付线, 0) = 1, True, False)
    g病人身份_黔南.中途结算 = Nvl(rsTemp!中途结帐, 0) = 1
    
    '进行反结算
    '交易特定输入数据：  交易类型|住院(门诊)号|单据号|操作员姓名|账户消费标志
    '交易特定输出数据:  费用总额|统筹支付|账户支付|现金支付|大病垫付|所有16项动态信息|交易流水号
    '交易类型定义如下：
    '   1正常结算 (出院结算)
    '   0住院中途结算
    '   -1反结算
    '   -2IC卡挂失后出院结算，本次结算（只针对住院）将所有费用转为现金支付，待到医保中心报销。
    '账户消费标志 0 不从账户消费（即标志为0，不下个人帐户）， 1  使用系统参数设置值（即标志为1，下个人帐户）。
    StrInput = "-1|"
    StrInput = StrInput & g病人身份_黔南.住院号 & "|"
    StrInput = StrInput & lng结帐ID & "|"
    StrInput = StrInput & gstrUserName & "|"
    StrInput = StrInput & IIf(g病人身份_黔南.下个人帐户, 1, 0)
    If 业务请求_黔南(结算, StrInput, strOutput) = False Then
        Exit Function
    End If
    strArr = Split(strOutput, "|")
    str动态信息 = ""
    '获取动态信息
    For i = 6 To UBound(strArr) - 1
        str动态信息 = str动态信息 & "|" & strArr(i)
    Next
    
    str动态信息 = Mid(str动态信息, 2)
    

    Dim objData As 结算数据
    With objData
        .费用总额 = Val(strArr(1))
        .统筹支付 = Val(strArr(2))
        .账户支付 = Val(strArr(3))
        .现金支付 = Val(strArr(4))
        .大病垫付 = Val(strArr(5))
        .动态信息 = str动态信息
        .交易流水号 = strArr(UBound(strArr))
    End With
            
    '确定是否需要写卡
    '写卡.
    '   如果病人账户余额不为零，根据结算函数返回的"账户支付"调用卡函数SaleTrans进行下帐操作。
    '   并且调用卡函数UpdateJrtclj, UpdateQfxzflj和卡函数UpdateDynInfo对卡动态信息进行更新操作。
    '   注意一定要调用卡函数UpdateHospStatus将病人的住院状态设置为2（出院状态），
    '   并调用卡函数AddHospTimes将病人的住院次数加一
    
    '??需加入写卡代码
    'bytType:1-SiCardBaseInfo社保卡基本信息
    '        2-SiCardDynaInfo社保卡动态信息
    '        3-SiCardAcctInfo社保卡帐户信息
    '        4-SiCardExtInfo社保卡扩展信息
    
    '进行写卡
    '设置动态属性
    If sCard_属性值(2, str动态信息, True) = False Then
        GoTo Err冲正:
    End If
    
    '写扩展信息
    '定点医院1|定点医院2|定点医院3|定点医院4|定点医院5|出院日期|住院状态（1—住院；2—出院）|就诊医院|医疗类别
    'bytType:1-SiCardBaseInfo社保卡基本信息
    '        2-SiCardDynaInfo社保卡动态信息
    '        3-SiCardAcctInfo社保卡帐户信息
    '        4-SiCardExtInfo社保卡扩展信息
    If g病人身份_黔南.中途结算 Then
    Else
        If sCard_属性值(4, Get扩展信息("1"), True) = False Then
            GoTo Err冲正:
        End If
    End If
    If sCard_SaveCard = False Then GoTo Err冲正:
    
    
    '帐户年度信息
    Call Get帐户信息(TYPE_黔南, rsTemp("病人ID"), Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & rsTemp("病人ID") & "," & TYPE_黔南 & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - rsTemp("个人帐户支付") & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")
    
    Dim dbl帐户余额 As Double
    If Save保险帐户_帐户余额(Nvl(rsTemp!病人ID, 0), str动态信息, dbl帐户余额) = False Then
        GoTo Err冲正:
    End If
        

   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(帐户增加累计),帐户累计支出_IN(帐户累计支出),累计进入统筹_IN(累计进入统筹_IN),累计统筹报销_IN(累计统筹报销),住院次数_IN(住院次数累计),起付线(无),封顶线_IN(无),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(无),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算时产生流水号),主页ID_IN,中途结帐_IN,备注_IN
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_黔南 & "," & rsTemp("病人ID") & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - rsTemp("个人帐户支付") & "," & cur进入统筹累计 & "," & cur统筹报销累计 & "," & int住院次数累计 & ",0,0," & Nvl(rsTemp!实际起付线, 0) & "," & _
        Nvl(rsTemp("发生费用金额"), 0) * -1 & "," & dbl帐户余额 & ",0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0,0," & _
        Nvl(rsTemp("个人帐户支付"), 0) * -1 & ",'" & objData.交易流水号 & "'," & rsTemp("主页ID") & "," & Nvl(rsTemp("中途结帐"), 0) & ",'" & Nvl(rsTemp!备注) & "'" & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "重庆医保")

    住院结算冲销_黔南 = True
    Exit Function
Err冲正:
    Call 业务请求_黔南(冲正交易, objData.交易流水号 & "|10|" & gstrUserName, strOutput)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function Get流水号(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng收费细目ID As Long, ByVal dbl数量 As Double, ByVal dbl单价 As Double, lng冲销ID As Long) As Variant
    '   随机取一条正常记录的流水号（用于负数记帐）,但需满足数量单价及费用一致.
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strArr  '0-交易流水号,1-审批编号,2-处方号,3-实际交易单价,4-实际等级....
    
    gstrSQL = " Select id, 摘要 From 住院费用记录" & _
              " Where 收费细目ID=[1] And 病人ID=[2] And 主页ID=[3]" & _
              " And 记录状态=1 And Nvl(是否上传,0)=1 and A.数次*nvl(A.付数,1)=[4] and Round(A.实收金额/(A.数次*nvl(A.付数,1)),4))=[5] And Nvl(实收金额,0)>0 And Rownum<2"
    
    DebugTool "进入获取流水号函数:GET流水号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取流水号", lng收费细目ID, lng病人ID, lng主页ID, dbl数量, dbl单价)
    If rsTemp.EOF Then
        strTemp = "||||||"
        lng冲销ID = 0
    Else
        strTemp = Nvl(rsTemp!摘要, "|") & "||||||"
        lng冲销ID = rsTemp!ID
    End If
    strArr = Split(strTemp, "|")
    Get流水号 = strArr
    DebugTool "结束获取流水号函数:GET流水号 返回值为:" & strTemp
End Function

Private Function 处方上传(ByVal lng记录性质 As Long, lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '处方明细上传
    '功能:上传新产生的记帐明细到医保中心
    '参数:  str单据号   NO
    '       int性质     记录性质
    '       lng病人ID  默认为0，表示传输整张单据，否则为单据中指定病人的。（主要是因为医嘱在保存记帐单时，是分病人在提交数据而不是一起提交）
    '返回:
    Dim rsTemp As New ADODB.Recordset, rs明细 As New ADODB.Recordset
    Dim StrInput As String, strOutput As String, strArr As Variant
    Dim str处方号 As String, str冲销明细流水号 As String, str审批编号 As String
    Dim lng病人ID As Long
    Dim bln单方 As Boolean
    
    处方上传 = False
    Err = 0
    On Error GoTo errHandle
    
   '读出该张单据的费用明细
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.主页ID,to_char(A.发生时间,'yyyy-mm-dd hh24:mi:ss') as 登记时间,Round(A.实收金额,4) 实收金额 " & _
              "         ,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         ,C.项目编码,C.附注,J.类别 as 收费类别,C.是否医保,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,E.规格,G.名称 剂型,M.医保号,M.就诊次数 " & _
              "  From 住院费用记录 A,收费类别 J,收费细目 B,保险帐户 M,(Select * From 保险支付项目 where 险类=" & TYPE_黔南 & ") C,病案主页 D,药品目录 E ,药品信息 F,药品剂型 G " & _
              "  where a.病人id=M.病人id  and M.险类=[1] and A.NO=[2] and A.记录性质=[3] and A.记录状态=1 And Nvl(A.是否上传,0)=0 " & _
              "        and A.收费类别=J.编码(+)  and A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=" & TYPE_黔南 & _
              "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID(+) " & _
              "        AND B.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) " & _
              "  Order by A.病人ID,A.发生时间"
              
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "处方明细上传", TYPE_黔南, str单据号, lng记录性质)
    Dim lng冲销ID As Long
    
    '先检查是否存在退单情况，如果存在，看有否对应的记录单.
    With rs明细
        '上传明细
        
        bln单方 = False
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 1 Then
            If InStr(1, "7中草药", Nvl(!收费类别)) <> 0 Then
                '单方的是自费
                bln单方 = True
            End If
        End If
        Do While Not .EOF
            '单价检查
            If 检查中心维护单价(Nvl(!病人ID, 0), Nvl(!收费细目ID, 0), Nvl(!价格, 0)) = False Then Exit Function
            .MoveNext
        Loop
    End With
    
    If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
    Dim strArr摘要
    
    '进行费用传输
    With rs明细
        Do Until .EOF
            gstrSQL = "Select * from 医保收费目录 where 类别=[1] and 编码=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "处方上传", CLng(Val(Nvl(!附注))), CStr(Nvl(!项目编码)))
            
            '上传明细记录
            '确定审批编号
            '修改记录:
            'modify by 2007-01-26 程池富
            '原因：所有项目都要进行项目审批结果查询

'            If Nvl(!是否医保, 1) = 0 Then
                StrInput = Nvl(!医保号) & "|"
                If Val(Nvl(!附注)) = 1 Then
                    '吕说:药品只能传医院编码,而诊疗只能等只能传医保编码
                    StrInput = StrInput & Nvl(!编码, "9000900099")
                Else
                    StrInput = StrInput & Nvl(!项目编码, "9000900099")
                End If
                If 业务请求_黔南(项目审批结果查询, StrInput, strOutput) = False Then
                    Exit Function
                End If
                '单价检查
                If Split(strOutput, "|")(1) <> "" Then
                    strArr = Split(strOutput, "|")
                    str审批编号 = strArr(1)
                Else
                    str审批编号 = ""
                End If
'            Else
'                str审批编号 = ""
'            End If
                
                lng病人ID = Nvl(!病人ID, 0)
                g病人身份_黔南.住院号 = lng病人ID & "_" & Nvl(!就诊次数, 0)
                '上传明细
                '交易特定输入数据：住院(门诊)号|处方号|审批编号|医院内码|医保编码|项目名称|
                '                  费用等级|费用类别|单价|数量|金额|单位|规格|剂型|开方日期|录入标志
                '交易特定输出数据:   该处方对应项目的实际单价|实际等级|交易流水号。
                str处方号 = rs明细!NO & "_" & lng记录性质 & "_" & lng记录状态
                StrInput = Nvl(!病人ID) & "_" & Nvl(!就诊次数, 0) & "|"
                StrInput = StrInput & str处方号 & "|"
                StrInput = StrInput & str审批编号 & "|"
                StrInput = StrInput & Nvl(!编码) & "|"
                StrInput = StrInput & IIf(bln单方, "9000900099", Nvl(!项目编码, "9000900099")) & "|"
                StrInput = StrInput & Nvl(!名称) & "|"
                
                If rsTemp.EOF Then
                    '吕说:中草药的需传1
                    If InStr(1, "7中草药", Nvl(!收费类别)) <> 0 Then
                        StrInput = StrInput & "1" & "|"
                    Else
                        StrInput = StrInput & "3" & "|"
                    End If
                    StrInput = StrInput & Split(Get费用类别(Nvl(!收费类别)), "-")(0) & "|"
                Else
                    '刘兴宏:2007/09/28加入,自费部分又是中药的,对应为甲类
                    '吕说:中草药的需传1
                    If InStr(1, "7中草药", Nvl(!收费类别)) And Nvl(!项目编码, "9000900099") = "9000900099" Then
                        StrInput = StrInput & "1" & "|"
                    Else
                        StrInput = StrInput & Nvl(rsTemp!收费等级) & "|"
                    End If
                    If IsNull(rsTemp!收费类别) Then
                        StrInput = StrInput & Split(Get费用类别(Nvl(!收费类别)), "-")(0) & "|"
                    Else
                        StrInput = StrInput & Nvl(rsTemp!收费类别) & "|"
                    End If
                End If
                
                StrInput = StrInput & Format(rs明细("价格"), "0.0000") & "|"
                StrInput = StrInput & Format(rs明细("数量"), "0.00") & "|"
                StrInput = StrInput & Format(rs明细("实收金额"), "#####0.00") & "|"         '金额
                
                StrInput = StrInput & ToVarchar(rs明细("计算单位"), 20) & "|"      '单位
                StrInput = StrInput & ToVarchar(rs明细("规格"), 14) & "|"
                StrInput = StrInput & ToVarchar(rs明细("剂型"), 20) & "|"
                StrInput = StrInput & Nvl(rs明细!登记时间) & "|"
            
                 '0 表示开始循环，2 表示结束循环：在结束循环后才提交
                If rs明细.AbsolutePosition = 1 Then
                    If rs明细.AbsolutePosition = rs明细.RecordCount Then
                        '只有一条记录时
                        StrInput = StrInput & "1"
                    Else
                        StrInput = StrInput & 0
                    End If
                ElseIf rs明细.AbsolutePosition = rs明细.RecordCount Then
                    StrInput = StrInput & 2
                Else
                    StrInput = StrInput & "1"
                End If
                
                If 业务请求_黔南(录入处方明细, StrInput, strOutput) = False Then
                    Exit Function
                End If
                
                '实际单价|实际等级|交易流水号
                strArr = Split(strOutput, "|")
                
                '摘要:交易流水号|审批编号||住院(门诊)号|处方号(门诊:门诊号；住院:单据号+记录性质)|实际交易单价|实际等级
                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strArr(3) & "|" & str审批编号 & "|" & g病人身份_黔南.住院号 & "|" & str处方号 & "|" & strArr(1) & "|" & strArr(2) & "')"
                zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            .MoveNext
        Loop
    End With
        
    处方上传 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function
Public Function 处方登记_黔南(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str病种代码 As String
    Dim dbl付数 As Double, dbl金额 As Double
    Dim StrInput As String, strOutput As String
    Dim str是否药品  As String
    Dim strArr
    
    Err = 0
    On Error GoTo errHand:
    
    
    处方登记_黔南 = False
'    If lng记录状态 = 1 Then
        '正常单据
        If 处方上传(lng记录性质, lng记录状态, str单据号) = False Then Exit Function
'    Else
'        '冲销单据
'        If 处方作废(lng记录性质, lng记录状态, str单据号) = False Then Exit Function
'    End If
        
    处方登记_黔南 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 处方作废(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:记帐处方作废,即记录状态=2的记录
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------
    Dim rs明细 As New ADODB.Recordset
    Dim rs原明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim str处方号 As String, StrInput As String, strOutput As String, str交易流水号 As String
    Dim strArr
    Dim lng病人ID As Long
    Dim str已被冲正交易 As String
    处方作废 = False
   
    Err = 0: On Error GoTo errHand:
          
    '检查该单据的原始单据是否存在负数
    
    gstrSQL = " Select 摘要,A.ID,a.收费细目id,A.序号,A.数次*nvl(A.付数,1) as 数量,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4) as 单价 " & _
              " From 住院费用记录 A,保险帐户 B " & _
              " where a.病人id=b.病人id and A.NO=[1] and A.记录性质=[2] and A.记录状态=3 and   Nvl(附加标志,0)<>9  order by A.病人id"
    Set rs原明细 = zlDatabase.OpenSQLRecord(gstrSQL, "处方明细上传", str单据号, lng记录性质)
    If rs原明细.EOF Then
        ShowMsgbox "该单据没有相应的明细记录,不能作废!"
        Exit Function
    End If
    
    gstrSQL = " Select * " & _
              " From 住院费用记录 A,保险帐户 b" & _
              " where a.病人id=b.病人id and A.NO=[1] and A.记录性质=[2] and A.记录状态=2 and Nvl(附加标志,0)<>9 AND nvl(a.是否上传,0)=0 "
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "处方明细上传", str单据号, lng记录性质)
    
    lng病人ID = 0
    '更新原单据的值
    With rs明细
        '摘要的值保存为:"交易流水号|审批编号|住院(门诊)号|处方号(门诊:门诊号；住院:单据号+记录性质)|实际交易单价|实际等级"
        Do While Not .EOF
            rs原明细.Filter = "序号=" & Nvl(!序号, 0) & "  and 收费细目id=" & Nvl(!收费细目ID, 0)
            If rs原明细.EOF Then
                ShowMsgbox "冲销时未找到相应的记录,冲销失败!"
                Exit Function
            End If
            strArr = Split(Nvl(rs原明细!摘要) & "|||||", "|")
            str交易流水号 = strArr(0)
            If str交易流水号 = "" Then
                ShowMsgbox "在原单中不存在交易流水号,不能继续！"
                Exit Function
            End If
            
            '更新上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(rs原明细!摘要) & "')"
            zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        str已被冲正交易 = ""
        
        Do While Not .EOF
            rs原明细.Filter = "序号=" & Nvl(!序号, 0) & "  and 收费细目id=" & Nvl(!收费细目ID, 0)
            '主要是记帐表时，需确定各病人的退方
            '"交易流水号|审批编号|住院(门诊)号|处方号(门诊:门诊号；住院:单据号+记录性质)|实际交易单价|实际等级"
            If lng病人ID <> Nvl(!病人ID, 0) Then
                '主要是记帐表时，需确定各病人的退方
                '"交易流水号|审批编号|住院(门诊)号|处方号(门诊:门诊号；住院:单据号+记录性质)|实际交易单价|实际等级"
                strArr = Split(Nvl(rs原明细!摘要) & "|||||", "|")
                g病人身份_黔南.住院号 = strArr(2)
                str处方号 = strArr(3)
                StrInput = g病人身份_黔南.住院号 & "|"
                StrInput = StrInput & str处方号
                If 业务请求_黔南(处方退方, StrInput, strOutput) = False Then Exit Function
                lng病人ID = Nvl(!病人ID, 0)
            End If
            .MoveNext
        Loop
    End With
    
 '   rs原明细.Filter = "数量<0 or 价格<0 "
 '   If rs原明细.EOF Then
        '退处方单,全单退
        '   交易特定输入数据：   住院(门诊)号|处方号
        '   交易特定输出数据:               退掉的处方明细的记录数量
'        strInput = g病人身份_黔南.住院号 & "|"交
'        strInput = strInput & str处方号
'        If 业务请求_黔南(处方退方, strInput, strOutput) = False Then Exit Function
'    Else
        '原单据存在负数冲帐,需调用冲正
'        With rs原明细
'            .Filter = 0
'            Do While Not .EOF
'                '交易特定输入数据： 被冲正交易流水号|被冲正交易类型代码|操作员代码
'                '交易特定输出数据:  未用
'                '摘要:"交易流水号|审批编号|住院(门诊)号|处方号(门诊:门诊号；住院:单据号+记录性质)|实际交易单价|实际等级|冲销标志"
'
'                '需查找该表正常单据是否已经被正常输入的处方单(负数)冲销.
'
'                strArr = Split(Nvl(!摘要, "|||||"), "|")
'                If Val(strArr(6)) = 1 Then
'                    '表明该记录已经被正常输入的负记录冲销不能再冲
'                Else
'                    strInput = strArr(0) & "|"
'                    strInput = Get交易代码(录入处方明细)
'                    If 业务请求_黔南(冲正交易, strInput, strOutput) = False Then Exit Function
'                End If
'                .MoveNext
'            Loop
'        End With
'    End If
    处方作废 = True
    Exit Function
errHand:
   If ErrCenter = 1 Then
        Resume
   End If
End Function
Private Function Read模拟数据(ByVal int业务类型 As 业务类型_黔南, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--功  能:通过该功能读取模拟数据,以便测试
    '--入参数:
    '--出参数:
    '--返  回:字串
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim STRNAME As String
    
    If int业务类型 = 读取卡内数据 Then
        strFile = App.Path & "\解析卡.txt"
    Else
        strFile = App.Path & "\模拟提交串.txt"
    End If
    
    
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    STRNAME = Get交易代码(int业务类型, True)
    
    Dim blnStart As Boolean
    Dim strArr
    
    Err = 0
    On Error GoTo errHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                If int业务类型 = 读取卡内数据 Then
                    strArr = Split(strText, vbTab)
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab
                            End If
                            strArr = Split(strText, vbTab)
                            
                            If Val(strArr(0)) = 1 Then
                                str = strArr(1)
                                Exit Do
                            End If
                        Else
                             If "<" & STRNAME & ">" = strText Then
                                 blnStart = True
                             End If
                        End If
                        If "</" & STRNAME & ">" = strText Then
                            Exit Do
                        End If
                End If
            Loop
            objText.Close
            strOutPutstring = str
    End If
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function

Private Function Get病人信息(ByVal lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strArr
    
    Get病人信息 = False
    'COMMENT ON COLUMN 保险帐户.病人ID   is '病人ID';
    'COMMENT ON COLUMN 保险帐户.险类     is '96';
    'COMMENT ON COLUMN 保险帐户.中心     is '0';
    'COMMENT ON COLUMN 保险帐户.卡号     is '卡号';
    'COMMENT ON COLUMN 保险帐户.医保号   is '医保证书号';
    'COMMENT ON COLUMN 保险帐户.密码     is '医疗类别';
    'COMMENT ON COLUMN 保险帐户.人员身份 is '目前未保存';
    'COMMENT ON COLUMN 保险帐户.单位编码 is '单位编码';
    'COMMENT ON COLUMN 保险帐户.顺序号   is '登记时的交易流水号';
    'COMMENT ON COLUMN 保险帐户.退休证号 is '获得通过的:审批编号|项目编号|项目名称';
    'COMMENT ON COLUMN 保险帐户.帐户余额 is '当前个人帐户余额';
    'COMMENT ON COLUMN 保险帐户.当前状态 is '0-门诊,1-在院';
    'COMMENT ON COLUMN 保险帐户.病种ID   is '病种ID，与保险病种的ID关联';
    'COMMENT ON COLUMN 保险帐户.在职     is '目前保存的值是1，无用处';
    'COMMENT ON COLUMN 保险帐户.年龄段   is '目前保存的是医保病人年龄';
    'COMMENT ON COLUMN 保险帐户.灰度级   is '审批类别';
    'COMMENT ON COLUMN 保险帐户.就诊时间 is '当前就诊的时间';
    'COMMENT ON COLUMN 保险帐户.就诊次数 is '病人当前就诊的次数,病人ID-就诊次数构成目前的门诊号';
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "select a.*,b.姓名,b.性别, b.年龄, b.出生日期, b.身份证号,b.工作单位 " & _
             " from 保险帐户 a,病人信息 b " & _
             " WHERE a.病人id=" & lng病人ID & " AND a.病人id=b.病人id and a.险类=" & TYPE_黔南
 
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人信息"
    
    With g病人身份_黔南
        .卡号 = Nvl(rsTemp!卡号)
        .医保证号 = Nvl(rsTemp!医保号)
        
        .姓名 = Nvl(rsTemp!姓名)
        .性别 = Nvl(rsTemp!性别)
        .年龄 = Nvl(rsTemp!年龄段, 0)
        .出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
        .单位编码 = Nvl(rsTemp!单位编码)
      
        strTemp = Nvl(rsTemp!工作单位)
        If InStr(1, strTemp, "(") <> 0 Then
            .单位名称 = Mid(strTemp, 1, InStr(1, strTemp, "(") - 1)
        Else
            .单位名称 = strTemp
        End If
        
        .医疗类别 = Nvl(rsTemp!密码)
        strArr = Split(Nvl(rsTemp!退休证号) & "||", "|")
        
        .审批编号 = strArr(0)
        .项目编码 = strArr(1)
        .项目名称 = strArr(2)
        .帐户余额 = Val(Nvl(rsTemp!帐户余额))
        .审批类别 = Nvl(rsTemp!灰度级)
        .住院号 = lng病人ID & "_" & Nvl(rsTemp!就诊次数, 0)
        .门诊号 = lng病人ID & "_" & Nvl(rsTemp!就诊次数, 0)
        .身份证号 = Nvl(rsTemp!身份证号)
        .病种ID = Nvl(rsTemp!病种ID, 0)
        
        '周玉强要求能在费用查询中进行虚拟结算
        .动态信息 = Nvl(rsTemp!动态信息)
        If .病种ID <> 0 Then
           gstrSQL = "Select 编码,名称 From 医保病种目录 where id=" & .病种ID
           OpenRecordset_黔南 rsTemp, "获取病种"
           
           If rsTemp.EOF Then
                .病种编码 = ""
                .病种名称 = ""
           Else
                .病种编码 = Nvl(rsTemp!编码)
                .病种名称 = Nvl(rsTemp!名称)
           End If
        Else
            .病种编码 = ""
            .病种名称 = ""
        End If
    End With
    Get病人信息 = True
Exit Function
errHand:
        DebugTool "获取病人信息失败" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & " 错误信息:" & Err.Description
End Function

Private Sub OpenRecordset_黔南(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional cnOracle As ADODB.Connection)
    '功能：打开记录集
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    If cnOracle Is Nothing Then
        rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_黔南, adOpenStatic, adLockReadOnly
    Else
        If cnOracle.State <> 1 Then
            rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_黔南, adOpenStatic, adLockReadOnly
        Else
            rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), cnOracle, adOpenStatic, adLockReadOnly
        End If
    End If
    Call SQLTest
End Sub


Public Function 住院虚拟结算_黔南(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String
    
    'rsExse:字符集
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    Dim cn上传 As New ADODB.Connection, rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim lng主页ID As Long
    Dim StrInput As String, strOutput   As String
    Dim strArr As Variant
    Dim cur个人帐户 As Double, cur统筹支付 As Double, cur大额统筹 As Double, cur公务员补助 As Double, cur发生费用 As Double
    Dim str总金额医院 As String, str总金额医保 As String, str冲销明细流水号 As String
    Dim str医生 As String, datCurr As Date, intMsg As Integer
    Dim str入院日期 As String, str出院日期 As String
    Dim intMouse As Integer
    
    Err = 0: On Error GoTo errHand:
    
    g病人身份_黔南.病人ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If
    If Get病人信息(lng病人ID) = False Then Exit Function
    
    If bln结帐处 Then
        Screen.MousePointer = 1
        If 身份标识_黔南(4, lng病人ID) = "" Then
            Screen.MousePointer = intMouse
            住院虚拟结算_黔南 = ""
            Exit Function
        End If
        Screen.MousePointer = intMouse
    Else
        g病人身份_黔南.下个人帐户 = True
    End If
    
    rsExse.MoveFirst
    datCurr = zlDatabase.Currentdate
    With g结算数据
        .病人ID = rsExse("病人ID")
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", CLng(rsExse("病人ID")))
        If IsNull(rsTemp("主页ID")) = True Then
            MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        lng主页ID = rsTemp("主页ID")
    End With

    Screen.MousePointer = vbHourglass
    
    '1.2 读出病人的入院时间
    gstrSQL = "" & _
        "   Select 入院日期,出院日期 " & _
        "   From 病案主页 where 病人ID=[1] and 主页ID=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", g结算数据.病人ID, lng主页ID)
    If IsNull(rsTemp("出院日期")) Then
        g病人身份_黔南.中途结算 = 1
        str出院日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    Else
        '表示该病人已经出院
        g病人身份_黔南.中途结算 = 0
        str出院日期 = Format(rsTemp!出院日期, "yyyy-mm-dd")
    End If
    '求和
    str入院日期 = Format(rsTemp!入院日期, "yyyy-mm-dd")
    
    g病人身份_黔南.费用总额 = 0
    Do While Not rsExse.EOF
        g病人身份_黔南.费用总额 = g病人身份_黔南.费用总额 + rsExse("金额")
        rsExse.MoveNext
    Loop
    g病人身份_黔南.费用总额 = Round(g病人身份_黔南.费用总额, 2)
    
    '补传明细
    If 补传住院明细记录(lng病人ID, lng主页ID) = False Then Exit Function
    
    '更新就诊信息
    gstrSQL = "" & _
         " select max(decode(A.诊断类型,1,b.编码||'~^||'||b.名称,null)) as 入院诊断,  " & _
         "        max(decode(A.诊断类型,1,null,b.编码)) as 确诊编码 " & _
         " from 诊断情况 A,疾病编码目录 b " & _
         " where a.疾病id=b.id and  a.诊断类型 in(1,2) and a.诊断次序=1 and a.病人id=" & lng病人ID & " and a.主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定诊断编码和名称"
    
    g病人身份_黔南.确诊诊断编码 = Nvl(rsTemp!确诊编码)
    
    If 更新就诊信息_黔南(2, strOutput, "", str出院日期, , IIf(bln结帐处, False, True)) = False Then Exit Function
    
    
    
    '返回参数:更新动态信息标志|所有16位动态信息（目前不处理.）
    
    '预结处理
    '交易特定输入数据：   住院（门诊）号|帐户消费标志
    '交易特定输出数据:   费用总额|统筹支付|账户支付|现金支付|大病垫付
    '3． 账户消费标志 0 不从账户消费（即标志为0，不下个人帐户）， 1  使用系统参数设置值（即标志为1，下个人帐户）
    
    Dim str结算方式  As String
    StrInput = g病人身份_黔南.住院号 & "|"
    StrInput = StrInput & IIf(g病人身份_黔南.下个人帐户, "1", "0")
    If 业务请求_黔南(预结算, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    strArr = Split(strOutput, "|")
    
    With 虚拟结算数据
        .费用总额 = Val(strArr(1))
        .统筹支付 = Val(strArr(2))
        .账户支付 = Val(strArr(3))
        .现金支付 = Val(strArr(4))
        .大病垫付 = Val(strArr(5))
        str结算方式 = "个人帐户;" & .账户支付 & ";0"   '不能修改个人帐户，因为结算时已经不再传金额到前置机了
        If .统筹支付 > 0 Then
            str结算方式 = str结算方式 & "|医保基金;" & .统筹支付 & ";0"
        End If
        If .大病垫付 > 0 Then
            str结算方式 = str结算方式 & "|大病垫付;" & .大病垫付 & ";0"
            If bln结帐处 = True Then MsgBox "该病人已进入大病统筹，请注意将[大病垫付：" & .大病垫付 & "]收取为现金！", vbInformation, gstrSysName
        End If
        If .费用总额 <> g病人身份_黔南.费用总额 Then
            ShowMsgbox "中心费用结算总额(" & .费用总额 & " ) 不等于医院实际发生费用总额(" & g病人身份_黔南.费用总额 & ")"
            Exit Function
        End If
    End With
    住院虚拟结算_黔南 = str结算方式
    g病人身份_黔南.病人ID = lng病人ID   '表示该病人已经进行了虚拟结算
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Get费用类别(ByVal str收费类别 As String) As String
    '获取费用类别
    Dim rsTemp As New ADODB.Recordset
    Dim str收费 As String
     str收费 = str收费类别
    If zlCommFun.ActualLen(str收费类别) = 1 Then
            gstrSQL = "Select * From 收费类别 where 编码='" & str收费类别 & "'"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取收费类别"
            
            If rsTemp.EOF Then
            Else
                str收费 = Nvl(rsTemp!类别)
            End If
    End If
    gstrSQL = "Select * From 保险参数 where 参数名='" & str收费 & "' and 险类=" & TYPE_黔南
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取类别"
    If rsTemp.EOF Then
        Get费用类别 = ""
    Else
        Get费用类别 = Nvl(rsTemp!参数值)
        
    End If
    If Get费用类别 = "" Then
        Get费用类别 = "-"
    End If
End Function
Private Function 补传住院明细记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '补传相关明细记录
    Dim cnTemp As New ADODB.Connection
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput  As String, strOutput As String
    Dim strArr, strArr摘要
    Dim lng冲销ID As Long
    Dim str冲销明细流水号 As String, str审批编号 As String, str处方号 As String
    Err = 0
    On Error GoTo errHand:
      
      
    补传住院明细记录 = False
    
    '读出未上传明细（排序，以便先上传正明细，再上传负明细）
    gstrSQL = "Select A.ID,A.NO,A.记录性质,J.类别 as 收费类别,A.记录状态,A.序号,A.病人ID,A.主页ID,to_char(A.发生时间,'yyyy-mm-dd hh24:mi:ss')  as 登记时间,Round(A.实收金额,4) 实收金额" & _
              "         ,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         ,C.项目编码,C.附注,C.是否医保,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,E.规格,G.名称 剂型 " & _
              "  From 住院费用记录 A,收费类别 J,收费细目 B, (Select * From 保险支付项目 where 险类=[1]) C,病案主页 D,药品目录 E ,药品信息 F,药品剂型 G " & _
              "  where A.病人ID=[2] and A.主页ID=[3] and A.记帐费用=1 and nvl(A.实收金额,0)<>0 and nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 " & _
              "        and A.收费类别=J.编码(+) and A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=[1]" & _
              "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID(+) " & _
              "        AND B.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) " & _
              "  Order by A.发生时间,A.记录性质,Decode(A.记录状态,2,2,1)"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", TYPE_黔南, lng病人ID, lng主页ID)
    
    Call DebugTool("打开新连接")
    Set cnTemp = GetNewConnection
    Call DebugTool("打开连接成功，开始检查明细数据的合法性。")
    
    '先检查是否存在退单情况，如果存在，看有否对应的记录单.
    With rs明细
        '上传明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '单价检查
            If 检查中心维护单价(Nvl(!病人ID, 0), Nvl(!收费细目ID, 0), Nvl(!价格, 0)) = False Then Exit Function
            .MoveNext
        Loop
    End With
    
    Call DebugTool("开始明细上传。")
    With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
           gstrSQL = "Select * from 医保收费目录 where 类别=[1] and 编码=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "处方上传", CLng(Val(Nvl(!附注))), CStr(Nvl(!项目编码)))
            If rsTemp.EOF Then
'                MsgBox "有项目未设置医保编码，不能上传明细!", vbInformation, gstrSysName
'                Exit Function
            End If
            'If Nvl(rs明细!记录状态, 0) Mod 3 = 1 Or Nvl(rs明细!记录状态, 0) Mod 3 = 0 Then
                Call DebugTool("上传正常记录。" & rs明细!NO)
                    
                    '正常处方单上传
                    '确定审批编号
                    If Nvl(!是否医保, 1) = 0 Then
                        StrInput = g病人身份_黔南.医保证号 & "|"
                        If Val(Nvl(!附注)) = 1 Then
                            '吕说:药品只能传医院编码,而诊疗只能等只能传医保编码
                            StrInput = StrInput & Nvl(!编码, "9000900099")
                        Else
                            StrInput = StrInput & Nvl(!项目编码, "9000900099")
                        End If
                        
                        If 业务请求_黔南(项目审批结果查询, StrInput, strOutput) = False Then
                            Exit Function
                        End If
                        strArr = Split(strOutput, "|")
                        str审批编号 = strArr(1)
                    Else
                        str审批编号 = ""
                    End If
                    '上传明细
                    '交易特定输入数据：住院(门诊)号|处方号|审批编号|医院内码|医保编码|项目名称|
                    '                  费用等级|费用类别|单价|数量|金额|单位|规格|剂型|开方日期|录入标志
                    '交易特定输出数据:   该处方对应项目的实际单价|实际等级|交易流水号。
                    str处方号 = rs明细!NO & "_" & Nvl(!记录性质) & "_" & Nvl(rs明细!记录状态)
                
                    StrInput = g病人身份_黔南.住院号 & "|"
                    StrInput = StrInput & str处方号 & "|"
                    StrInput = StrInput & str审批编号 & "|"
                    StrInput = StrInput & Nvl(!编码) & "|"
                    StrInput = StrInput & Nvl(!项目编码, "9000900099") & "|"
                    StrInput = StrInput & Nvl(!名称) & "|"
        
                    'Modify by 程池富 20090608 中草药退方错误
                    If rsTemp.EOF Then
                        '吕说:中草药的需传1
                        If InStr(1, "7中草药", Nvl(!收费类别)) <> 0 Then
                            StrInput = StrInput & "1" & "|"
                        Else
                            StrInput = StrInput & "3" & "|"
                        End If
                        StrInput = StrInput & Split(Get费用类别(Nvl(!收费类别)), "-")(0) & "|"
                    Else
                        '刘兴宏:2007/09/28加入,自费部分又是中药的,对应为甲类
                        '吕说:中草药的需传1
                        If InStr(1, "7中草药", Nvl(!收费类别)) And Nvl(!项目编码, "9000900099") = "9000900099" Then
                            StrInput = StrInput & "1" & "|"
                        Else
                            StrInput = StrInput & Nvl(rsTemp!收费等级) & "|"
                        End If
                        If IsNull(rsTemp!收费类别) Then
                            StrInput = StrInput & Split(Get费用类别(Nvl(!收费类别)), "-")(0) & "|"
                        Else
                            StrInput = StrInput & Nvl(rsTemp!收费类别) & "|"
                        End If
                    End If
                    
                    StrInput = StrInput & Format(rs明细("价格"), "0.0000") & "|"
                    StrInput = StrInput & Format(rs明细("数量"), "0.00") & "|"
                    StrInput = StrInput & Format(rs明细("实收金额"), "#####0.0000") & "|"         '金额
                
                    StrInput = StrInput & ToVarchar(rs明细("计算单位"), 20) & "|"      '单位
                    StrInput = StrInput & ToVarchar(rs明细("规格"), 14) & "|"
                    StrInput = StrInput & ToVarchar(rs明细("剂型"), 20) & "|"
                    StrInput = StrInput & Nvl(rs明细!登记时间) & "|"
                    StrInput = StrInput & 1
                    
                    If 业务请求_黔南(录入处方明细, StrInput, strOutput) = False Then
                        Exit Function
                    End If
                    '实际单价|实际等级|交易流水号
                    strArr = Split(strOutput, "|")
                    '摘要:交易流水号|审批编号||住院(门诊)号|处方号(门诊:门诊号；住院:单据号+记录性质)|实际交易单价|实际等级
                    gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strArr(3) & "|" & str审批编号 & "|" & g病人身份_黔南.住院号 & "|" & str处方号 & "|" & strArr(1) & "|" & strArr(2) & "')"
                    cnTemp.Execute gstrSQL, , adCmdStoredProc
'                End If
'            Else
'                '是上传被冲销的记录,则冲正
'                'strArr摘要 :0交易流水号,1审批编号,2住院(门诊)号,3处方号,4实际交易单价,5实际等级,6-是否被冲销"
'                strArr摘要 = Get原单据摘要(Nvl(!NO), Nvl(!序号), Nvl(!记录性质, 0))
'                '需冲正相应的流水号数据
'                '交易特定输入数据：  被冲正交易流水号|被冲正交易类型代码|操作员代码
'                '交易特定输出数据:   未用
'                Call DebugTool("上传冲销记录。str原接要流水号为" & strArr摘要(0))
'                If Val(strArr摘要(6)) = 1 Then
'                        '表明该记录已经被正常输入的负记录冲销不能再冲
'                Else
'                    StrInput = strArr摘要(0) & "|"
'                    StrInput = StrInput & Get交易代码(录入处方明细) & "|"
'                    StrInput = StrInput & ToVarchar(Nvl(!操作员姓名, gstrUserName), 20)
'                    If 业务请求_黔南(冲正交易, StrInput, strOutput) = False Then Exit Function
'                End If
'                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strArr摘要(0) & "|" & strArr摘要(1) & "|" & strArr摘要(2) & "|" & strArr摘要(3) & "|" & strArr摘要(4) & "|" & strArr摘要(5) & "|" & 0 & "')"
'                cnTemp.Execute gstrSQL, , adCmdStoredProc
'            End If
            .MoveNext
        Loop
    End With
    Call DebugTool("明细上传成功!")
    
    补传住院明细记录 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call DebugTool("明细上传失败!")
End Function
Private Function Get原单据摘要(ByVal strNO As String, ByVal int序号 As Integer, ByVal int性质 As Integer) As Variant
    '根据指定的值，获取摘要的相关信息
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
        
    
    gstrSQL = " Select 摘要 From 住院费用记录" & _
              " Where NO=[1] And 序号=[2]" & _
              " And 记录性质=[3] And 记录状态=3"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取原始处方明细的流水号", strNO, int序号, int性质)
    
    If Not rsTemp.EOF Then
        strTemp = Nvl(rsTemp!摘要) & "|||||||"
    Else
        strTemp = "|||||||"
    End If
    Get原单据摘要 = Split(strTemp, "|")
End Function

'----200410刘兴宏加入
Public Function 医保设置_黔南() As Boolean
    医保设置_黔南 = frmSet黔南.参数设置
    
End Function
'
'Public Function 下载服务项目目录_黔南(ByVal bytType As Byte, ByVal objProgss As Object) As Boolean
'    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:下载服务项目目录
'    '参数:bytType-1-药品,2-诊疗,3-服务,4-费用类别,5-病种目录
'    '返回:下载成功,返回true,否则返回False
'    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSql As String
'    Dim rsTemp As New ADODB.Recordset
'    Dim strDate As String, strInput As String, strOutput As String
'    Dim lngCount As Long
'    Dim i As Long
'    Dim strArr
'    Dim strTitle As String
'
'    下载服务项目目录_黔南 = False
'    strTitle = Switch(bytType = 1, "药品", bytType = 2, "诊疗项目", bytType = 3, "服务设施", bytType = 4, "费用类别", True, "病种目录取")
'
'    Err = 0
'    On Error GoTo ErrHand:
'    strSql = "" & _
'        "   Select to_char(Max(变更时间),'yyyy-mm-dd hh24:mi:ss')  as 变更时间 " & _
'        "   From 医保收费目录 " & _
'        "   where 类别=" & bytType
'    zlDatabase.OpenRecordset rsTemp, strSql, "获取最大变更时间"
'
'    strDate = Nvl(rsTemp!变更时间)
'    strDate = IIf(strDate = "", "1977-01-01 00:00:00", strDate)
'
'    If Not objProgss Is Nothing Then
'    Else
'        zlCommFun.ShowFlash "正在下载《" & strTitle & "》数据,请等待..."
'    End If
'    '预处理
'    strInput = bytType & "|" & strDate
'    If 业务请求_黔南(收费目录下载预处理, strInput, strOutput) = False Then Exit Function
'    strArr = Split(strOutput, "|")
'    lngCount = Val(strArr(1))
'
'    If Not objProgss Is Nothing Then
'        objProgss.Max = IIf(lngCount = 0, 1, lngCount)
'        objProgss.Min = 1
'        objProgss.Value = 1
'    End If
'
'   For i = 1 To lngCount
'        '正试下载
'        If 业务请求_黔南(收费目录下载处理, strInput, strOutput) = False Then Exit Function
'        strArr = Split(strOutput, "|")
'        '更新收费目录
'
'        '过程:类别,编码,名称,英文名称,收费类别,收费等级,费用等级,拼音码,单位,单价,规格,备注,变更时间,可维护标志,支付标准
'        strSql = "ZL_医保收费目录_UPDATE("
'        strSql = strSql & bytType & ",'"
'        strSql = strSql & strArr(1) & "','" '编码
'        strSql = strSql & strArr(2) & "','" '名称
'        Select Case bytType
'        Case 1
'            strSql = strSql & strArr(3) & "','" '英文名称
'            strSql = strSql & strArr(4) & "','" '收费类别
'            strSql = strSql & strArr(5) & "','" '费用等级
'            strSql = strSql & strArr(6) & "','" '拼音码
'            strSql = strSql & strArr(7) & "','" '单位
'            strSql = strSql & strArr(8) & "','" '单价
'            strSql = strSql & strArr(9) & "','" '剂型
'            strSql = strSql & strArr(10) & "','" '规格
'            strSql = strSql & strArr(11) & "',to_date('" '备注
'            strSql = strSql & strArr(12) & "','yyyy-mm-dd hh24:mi:ss'),'"  '变更时间
'            strSql = strSql & strArr(13) & "','"     '可维护标志
'            strSql = strSql & "" & "')" '支付标准
'        Case 2
'            strSql = strSql & "" & "','" '英文名称
'            strSql = strSql & strArr(3) & "','" '收费类别
'            strSql = strSql & "" & "','" '费用等级
'            strSql = strSql & strArr(4) & "','" '拼音码
'            strSql = strSql & strArr(5) & "','" '单位
'            strSql = strSql & strArr(6) & "','" '单价
'            strSql = strSql & "" & "','" '剂型
'            strSql = strSql & "" & "','" '规格
'            strSql = strSql & strArr(7) & "',to_date('" '备注
'            strSql = strSql & strArr(8) & "','yyyy-mm-dd hh24:mi:ss'),'"  '变更时间
'            strSql = strSql & strArr(9) & "','"     '可维护标志
'            strSql = strSql & "" & "')" '支付标准
'        Case 3
'            strSql = strSql & "" & "','" '英文名称
'            strSql = strSql & strArr(3) & "','" '收费类别
'            strSql = strSql & "" & "','" '费用等级
'            strSql = strSql & strArr(6) & "','" '拼音码
'            strSql = strSql & "" & "','" '单位
'            strSql = strSql & strArr(4) & "','" '单价
'            strSql = strSql & "" & "','" '剂型
'            strSql = strSql & "" & "','" '规格
'            strSql = strSql & "" & "',to_date('" '备注
'            strSql = strSql & strArr(7) & "','yyyy-mm-dd hh24:mi:ss'),'"  '变更时间
'            strSql = strSql & "" & "',"     '可维护标志
'            strSql = strSql & strArr(5) & "')" '支付标准
'        Case 4
'            ' 费用类别编码|费用类别名称
'            strSql = "ZL_医保收费类别_UPDATE("
'
'            strSql = strSql & strArr(1) & "','" '编码
'            strSql = strSql & strArr(2) & "')" '名称
'        Case Else
'            '病种编码|病种名称|拼音码|变更日期
'            strSql = "ZL_医保病种目录_UPDATE("
'            strSql = strSql & strArr(1) & "','" '编码
'            strSql = strSql & strArr(2) & "','" '名称
'            strSql = strSql & strArr(3) & "',to_date('" '助记码
'            strSql = strSql & strArr(4) & "','yyyy-mm-dd hh24:mi:ss')" '变更时间
'        End Select
'        gcnOracle_黔南.Execute strSql, , adCmdStoredProc
'        If Not objProgss Is Nothing Then
'            objProgss.Value = i
'        Else
'            zlCommFun.ShowFlash "正在下载《" & strTitle & "》数据,已下载" & i & "/" & lngCount & ""
'        End If
'   Next
'   下载服务项目目录_黔南 = True
'   Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'End Function




Public Function 下载服务项目目录_黔南(ByVal bytType As Byte, ByVal objProgss As Object) As Boolean
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:下载服务项目目录
    '参数:bytType-1-药品,2-诊疗,3-服务,4-费用类别,5-病种目录
    '返回:下载成功,返回true,否则返回False
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Dim strDate As String, StrInput As String, strOutput As String
    Dim lngCount As Long
    Dim i As Long
    Dim strArr
    Dim strTitle As String
    
    下载服务项目目录_黔南 = False
    

    If gcnOracle_中心 Is Nothing Then
        If Open中间库 = False Then Exit Function
    End If
    If gcnOracle_中心.State <> 1 Then
        If Open中间库 = False Then Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand:
    
    strSQL = "" & _
        "   Select to_char(Max(变更时间),'yyyy-mm-dd hh24:mi:ss')  as 变更时间 " & _
        "   From 医保收费目录 " & _
        "   where 类别=" & bytType
    zlDatabase.OpenRecordset rsTemp, strSQL, "获取最大变更时间"
    
    strDate = Nvl(rsTemp!变更时间)
    strDate = IIf(strDate = "", "1977-01-01 00:00:00", strDate)
       
    If Not objProgss Is Nothing Then
    Else
        zlCommFun.ShowFlash "正在下载《" & strTitle & "》数据,请等待..."
    End If
    
    Select Case bytType
    Case 1      '药品
        strTitle = "药品"
        gstrSQL = "Select * From medicine_info where AAE035>to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')"
    Case 2      '诊疗项目
        strTitle = "诊疗项目"
        gstrSQL = "Select * From examine_info where AAE035>to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')"
    Case 3      '服务设施
        strTitle = "服务设施"
        gstrSQL = "Select * From equipment_info where AAE035>to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')"
    
    Case 4      '费用类别
        strTitle = "费用类别"
        gstrSQL = "Select * From CHARGETYPE_INFO "
    Case 5      '病种目录
        strTitle = "病种目录"
        gstrSQL = "Select * From illness_info where AAE035>to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')"
    End Select
    
    
    OpenRecordset_黔南 rsData, "获取" & strTitle & "数据", , gcnOracle_中心
            
    If Not objProgss Is Nothing Then
        objProgss.Max = IIf(rsData.RecordCount = 0, 1, rsData.RecordCount) + 1
        objProgss.Min = 1
        objProgss.Value = 1
    End If
    i = 1
    With rsData
        Do While Not .EOF
            
            strSQL = "ZL_医保收费目录_UPDATE("
            strSQL = strSQL & bytType & ","
            Select Case bytType
            Case 1
                strSQL = strSQL & "'" & strProcess(!AKA060) & "',"    'AKA060  VARCHAR2(20)                   药品编码
                strSQL = strSQL & "'" & strProcess(!AKA061) & "',"   'AKA061 VARCHAR2(50)  Y                中文名称
                strSQL = strSQL & "'" & strProcess(!AKA062) & "',"  'AKA062 VARCHAR2(50)  Y                英文名称
                strSQL = strSQL & "'" & strProcess(!AKA063) & "'," 'AKA063 VARCHAR2(3)   Y                收费类别
                strSQL = strSQL & "'" & strProcess(!AKA065) & "'," 'AKA065 VARCHAR2(3)   Y                收费项目等级
                strSQL = strSQL & "'" & strProcess(!AKA066) & "',"  'AKA066 VARCHAR2(14)  Y                助记码
                strSQL = strSQL & "'" & strProcess(!AKA067) & "',"  'AKA067 VARCHAR2(20)  Y                单位
                strSQL = strSQL & "" & Val(strProcess(!AKA068)) & ","  'AKA068 NUMBER(8,2)   Y                标准价格
                strSQL = strSQL & "'" & strProcess(!AKA070) & "',"   'AKA070 VARCHAR2(50)  Y                剂型
                strSQL = strSQL & "'" & strProcess(!AKA074) & "',"    'AKA074 VARCHAR2(50)  Y                规格
                strSQL = strSQL & "'" & strProcess(!AAE013) & "',"      'AAE013 VARCHAR2(100) Y                备注
                'AAE035 DATE          Y                变更日期
                strSQL = strSQL & "to_date('" & Format(!AAE035, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                strSQL = strSQL & "'" & strProcess(!AAA104) & "',"     'AAA104 VARCHAR2(3)   Y                代码可维护标志
                strSQL = strSQL & "'" & "" & "',"  'AAA104 VARCHAR2(3)   Y                支付标准
                
                strSQL = strSQL & "'" & strProcess(!AKA305) & "',"      'AKA305 VARCHAR2(3)   Y                药品种类
                strSQL = strSQL & "'" & strProcess(!AKA064) & "',"      'AKA064 VARCHAR2(3)   Y                处方药标志
                strSQL = strSQL & "" & Val(strProcess(!AKA069)) & ","       'AKA069 NUMBER(5,4)   Y                自付比例
                strSQL = strSQL & "" & Val(strProcess(!AKA071)) & ","      'AKA071 NUMBER(5,2)   Y                每次用量
                strSQL = strSQL & "'" & strProcess(!AKA072) & "',"      'AKA072 VARCHAR2(20)  Y                使用频次
                strSQL = strSQL & "'" & strProcess(!AKA073) & "',"      'AKA073 VARCHAR2(50)  Y                用法
                strSQL = strSQL & "'" & strProcess(!AKA030) & "',"      'AKA030 VARCHAR2(3)   Y                当前使用标志
                strSQL = strSQL & "'" & strProcess(!AAB034) & "',"      'AAB034 VARCHAR2(14)  Y                社保经办机构
                '    医院等级_IN IN 医保收费目录.医院等级%TYPE)
                strSQL = strSQL & "NUll,"
                '   住院自费比例_IN IN 医保收费目录.住院自费比例%TYPE,
                strSQL = strSQL & "NUll,"
                '   参保人员_IN IN 医保收费目录.参保人员%TYPE
                strSQL = strSQL & "NUll)"
            
            Case 2
    
                'AKA090 VARCHAR2(20)                   项目编码
                '编码_IN IN 医保收费目录.编码%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA090) & "',"
                'AKA091 VARCHAR2(50)  Y                项目名称
                '名称_IN IN 医保收费目录.名称%TYPE,
                strSQL = strSQL & "'" & Replace(strProcess(!AKA091), "'", "‘") & "',"
                '    英文名称_IN IN 医保收费目录.英文名称%TYPE,
                strSQL = strSQL & "'" & "" & "',"
                'AKA063 VARCHAR2(3)   Y                收费类别
                '    收费类别_IN IN 医保收费目录.收费类别%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA063) & "',"
                'AKA065 VARCHAR2(3)   Y                收费项目等级
                '    收费等级_IN IN 医保收费目录.收费等级%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA065) & "',"
                'AKA066 VARCHAR2(14)  Y                助记码
                '    助记码_IN IN 医保收费目录.助记码%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA066) & "',"
                'AKA067 VARCHAR2(20)  Y                单位
                '    单位_IN IN 医保收费目录.单位%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA067) & "',"
                'AKA068 NUMBER(8,2)   Y                标准价格
                '    标准价格_IN IN 医保收费目录.标准价格%TYPE,
                strSQL = strSQL & "" & Val(strProcess(!AKA068)) & ","
                '    剂型_IN IN 医保收费目录.剂型%TYPE,
                strSQL = strSQL & "NUll,"
                '    规格_IN IN 医保收费目录.规格%TYPE,
                strSQL = strSQL & "NUll,"
                'AAE013 VARCHAR2(100) Y                备注
                '    备注_IN IN 医保收费目录.备注%TYPE,
                strSQL = strSQL & "'" & strProcess(!AAE013) & "',"
                'AAE035 DATE          Y                变更日期
                '    变更时间_IN IN 医保收费目录.变更时间%TYPE,
                strSQL = strSQL & "to_date('" & Format(!AAE035, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                'AAA104 VARCHAR2(3)   Y                代码可维护标志
                '    维护标志_IN IN 医保收费目录.维护标志%TYPE,
                strSQL = strSQL & "'" & strProcess(!AAA104) & "',"
                '    支付标准_IN IN 医保收费目录.支付标准%TYPE,
                strSQL = strSQL & "NUll,"
                '    药品种类_IN IN 医保收费目录.药品种类%TYPE,
                strSQL = strSQL & "NUll,"
                '    处方药标志_IN IN 医保收费目录.处方药标志%TYPE,
                strSQL = strSQL & "NUll,"
                'AKA069 NUMBER(5,4)   Y                自付比例
                '    自付比例_IN IN 医保收费目录.自付比例%TYPE,
                strSQL = strSQL & "" & Val(strProcess(!AKA069)) & ","
                '    每次用量_IN IN 医保收费目录.每次用量%TYPE,
                strSQL = strSQL & "NUll,"
                '    使用频次_IN IN 医保收费目录.使用频次%TYPE,
                strSQL = strSQL & "NUll,"
                '    用法_IN IN 医保收费目录.用法%TYPE,
                strSQL = strSQL & "NUll,"
                'AKA030 VARCHAR2(3)   Y                当前使用标志
                '    当前使用标志_IN IN 医保收费目录.当前使用标志%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA030) & "',"
                'AAB034 VARCHAR2(14)  Y                社保经办机构
                '    社保经办机构_IN IN 医保收费目录.社保经办机构%TYPE
                strSQL = strSQL & "'" & strProcess(!AAB034) & "',"
                'AKA101 VARCHAR2(3)                    医院等级
                '    医院等级_IN IN 医保收费目录.医院等级%TYPE)
                strSQL = strSQL & "'" & strProcess(!AKA101) & "',"
                'ZKA001 NUMBER(5,4)   Y                住院自付比例
                '   住院自费比例_IN IN 医保收费目录.住院自费比例%TYPE,
                strSQL = strSQL & "" & Val(strProcess(!ZKA001)) & ","
                'ZKA004 VARCHAR2(3)   Y                0-所有参保人员，1-普通职工，2-特殊人员
                '   参保人员_IN IN 医保收费目录.参保人员%TYPE
                strSQL = strSQL & "'" & strProcess(!ZKA004) & "')"
                
            Case 3
                'AKA100 VARCHAR2(20)                  医疗服务设施编码
                '编码_IN IN 医保收费目录.编码%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA100) & "',"
                'AKA102 VARCHAR2(50) Y                服务设施名称
                '名称_IN IN 医保收费目录.名称%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA102) & "',"
                '英文名称_IN IN 医保收费目录.英文名称%TYPE,
                strSQL = strSQL & "NULL,"
                'AKA063 VARCHAR2(3)  Y                收费类别
                '收费类别_IN IN 医保收费目录.收费类别%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA063) & "',"
                'AKA103 VARCHAR2(3)  Y                病床等级
                '收费等级_IN IN 医保收费目录.收费等级%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA103) & "',"
                'AKA066 VARCHAR2(14) Y                助记码
                '助记码_IN IN 医保收费目录.助记码%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA066) & "',"
                '单位_IN IN 医保收费目录.单位%TYPE,
                strSQL = strSQL & "NULL,"
                'AKA068 NUMBER(8,2)  Y                标准价格
                '标准价格_IN IN 医保收费目录.标准价格%TYPE,
                strSQL = strSQL & "" & Val(strProcess(!AKA068)) & ","
                '剂型_IN IN 医保收费目录.剂型%TYPE,
                strSQL = strSQL & "NULL,"
                '规格_IN IN 医保收费目录.规格%TYPE,
                strSQL = strSQL & "NULL,"
                '备注_IN IN 医保收费目录.备注%TYPE,
                strSQL = strSQL & "NULL,"
                'AAE035 DATE         Y                变更日期
                '变更时间_IN IN 医保收费目录.变更时间%TYPE,
                strSQL = strSQL & "to_date('" & Format(!AAE035, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                '维护标志_IN IN 医保收费目录.维护标志%TYPE,
                strSQL = strSQL & "NULL,"
                'AKA104 NUMBER(8,2)  Y                基础支付标准
                '支付标准_IN IN 医保收费目录.支付标准%TYPE,
                '
                strSQL = strSQL & "" & Val(strProcess(!AKA104)) & ","
                '药品种类_IN IN 医保收费目录.药品种类%TYPE,
                strSQL = strSQL & "NULL,"
                '处方药标志_IN IN 医保收费目录.处方药标志%TYPE,
                strSQL = strSQL & "NULL,"
                '自付比例_IN IN 医保收费目录.自付比例%TYPE,
                strSQL = strSQL & "NULL,"
                '每次用量_IN IN 医保收费目录.每次用量%TYPE,
                strSQL = strSQL & "NULL,"
                '使用频次_IN IN 医保收费目录.使用频次%TYPE,
                strSQL = strSQL & "NULL,"
                '用法_IN IN 医保收费目录.用法%TYPE,
                strSQL = strSQL & "NULL,"
                'AKA030 VARCHAR2(3)  Y                当前使用标志
                '当前使用标志_IN IN 医保收费目录.当前使用标志%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA030) & "',"
                'AAB034 VARCHAR2(14) Y                社保经办机构
                '社保经办机构_IN IN 医保收费目录.社保经办机构%TYPE,
                strSQL = strSQL & "'" & strProcess(!AAB034) & "',"
                'AKA101 VARCHAR2(3)                   医院等级
                '医院等级_IN IN 医保收费目录.医院等级%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA101) & "',"
                '住院自费比例_IN IN 医保收费目录.住院自费比例%TYPE,
                strSQL = strSQL & "NULL,"
                '参保人员_IN IN 医保收费目录.参保人员%TYPE
                strSQL = strSQL & "NULL)"
            Case 4
                ' 费用类别编码|费用类别名称
                strSQL = "ZL_医保收费类别_UPDATE("
                '编码_IN IN 医保收费目录.编码%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA063) & "',"
                '名称_IN IN 医保收费目录.名称%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA110) & "',"
                '    大类编码_IN IN 医保收费类别.大类编码%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA111) & "',"
                '    大类名称_IN IN 医保收费类别.大类名称%TYPE
                strSQL = strSQL & "'" & strProcess(!AKA112) & "')"
            
            Case Else
                strSQL = "ZL_医保病种目录_UPDATE("
                'AKA120 VARCHAR2(20)                  病种编码
                '    编码_IN IN 医保病种目录.编码%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA120) & "',"
                'AKA121 VARCHAR2(50) Y                病种名称
                '    名称_IN IN 医保病种目录.名称%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA121) & "',"
                'AKA066 VARCHAR2(14) Y                助记码
                '    助记码_IN IN 医保病种目录.助记码%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA066) & "',"
                'AKA122 VARCHAR2(3)  Y                病种分类
                '    病种分类_IN IN 医保病种目录.病种分类%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA122) & "',"
                'AKA124 NUMBER(6,4)  Y                病种自付比例
                '    自付比例_IN IN 医保病种目录.自付比例%TYPE,
                strSQL = strSQL & "" & Val(strProcess(!AKA124)) & ","
                'AKA030 VARCHAR2(3)  Y                当前使用标志
                '    当前使用标志_IN IN 医保病种目录.当前使用标志%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA030) & "',"
                'AKA125 NUMBER(8,2)  Y
                '    AKA125_IN IN 医保病种目录.AKA125%TYPE,
                strSQL = strSQL & "" & Val(strProcess(!AKA125)) & ","
                'AKA126 VARCHAR2(3)  Y                统计码
                '    统计码_IN IN 医保病种目录.统计码%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA126) & "',"
                'AKA128 VARCHAR2(14) Y                病种流水号
                '    病种流水号_IN IN 医保病种目录.病种流水号%TYPE,
                strSQL = strSQL & "'" & strProcess(!AKA128) & "',"
                'AAE035 DATE         Y                变更日期
                '    变更时间_IN IN 医保病种目录.变更时间%TYPE
                strSQL = strSQL & "to_date('" & Format(!AAE035, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))"
            End Select
            gcnOracle_黔南.Execute strSQL, , adCmdStoredProc
            If Not objProgss Is Nothing Then
                objProgss.Value = i
            Else
                zlCommFun.ShowFlash "正在下载《" & strTitle & "》数据,已下载" & i & "/" & lngCount & ""
            End If
            i = i + 1
            .MoveNext
        Loop
    End With
   下载服务项目目录_黔南 = True
   Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetItemInfo_黔南(ByVal lngPatiID As Long, ByVal lngItemID As Long, Optional ByVal str摘要 As String, Optional intType As Integer = 0) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取大连病人的相关提示信息
    '--入参数:
    '--出参数:
    '--返  回:提示串
    '-----------------------------------------------------------------------------------------------------------
    Dim strMsgInfor As String
    Dim str原摘要 As String, StrInput As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim str项目编码 As String
    Dim str医院内码 As String
    Dim bln药品 As Boolean
    
    str原摘要 = str摘要
    
    


    gstrSQL = "Select a.*,b.编码 from 保险支付项目 a,收费细目 b where a.收费细目id=b.id and  nvl(a.是否医保,0)=0 and 险类=" & TYPE_黔南 & " and a.收费细目iD=" & lngItemID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取是否特殊治疗，特殊检查和贵重药品!"
    If rsTemp.EOF Then
        Exit Function
    End If
    
    str项目编码 = Nvl(rsTemp!项目编码, "9000900099")
    str医院内码 = Nvl(rsTemp!编码, "9000900099")
    bln药品 = Val(Nvl(rsTemp!附注)) = 1
    
    gstrSQL = "Select 医保号 From 保险帐户 where 病人id=" & lngPatiID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取医保证号"
    If rsTemp.EOF Then
        ShowMsgbox "该病人不是医保病人!"
        Exit Function
    End If
    '检查是否审批
    '获取经过审批的项目编号
    StrInput = Nvl(rsTemp!医保号, 0) & "|"
    If bln药品 Then
    '吕说:药品只能传医院编码,而诊疗只能等只能传医保编码
        StrInput = StrInput & str医院内码
    Else
        StrInput = StrInput & str项目编码
    End If
    
    If 业务请求_黔南(项目审批结果查询, StrInput, strOutput) = False Then
        ShowMsgbox "该项目是否特检或检治或贵重药品，但未经过审批！"
        Exit Function
    End If
    
    '  执行成功时为审批编号(=0时),执行失败时（=100输出为空，为没有查到对应审批信息。）失败原因描述（< 0时，查询失败，返回失败原因描述）
    If Split(strOutput, "|")(1) = "" Then
        ShowMsgbox "注意:" & vbCrLf & "   该收费细目未能过审批,将以全自费处理！"
    End If
End Function
Public Function 获取参保人员信息_黔南() As Boolean
    '获取参保人员信息
    Dim StrInput As String
    Dim strOutput As String
    Dim strArr
    获取参保人员信息_黔南 = False
    
    Err = 0
    On Error GoTo errHand:
    
    '＊从卡上获取医保证号
    If 解析卡_黔南() = False Then Exit Function
    
    '＊根据医保证号获取参保信息
    StrInput = g病人身份_黔南.医保证号
    If 业务请求_黔南(获得参保人员信息, StrInput, strOutput) = False Then
        
        Exit Function
    End If
    If strOutput = "" Then
        ShowMsgbox "在获取参保人员信息时，接口返回了空值!"
        Exit Function
    End If
    strArr = Split(strOutput, "|")
    '返回: 姓名|性别|身份证|出生日期|人员类别编码|人员类别名称|单位名称|单位编码|统筹区号
    '   统筹区号：是东软公司吕智超要求要在后加入的参数
    With g病人身份_黔南
        
        .姓名 = strArr(1)
        .性别 = strArr(2)
        .身份证号 = strArr(3)
        .出生日期 = strArr(4)
        .类别编码 = strArr(5)
        .类别名称 = strArr(6)
        .单位名称 = strArr(7)
        .单位编码 = strArr(8)
        .统筹区号 = strArr(9)
        .年龄 = Get年龄(.出生日期)
    End With
    获取参保人员信息_黔南 = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function

Private Function Get年龄(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as 年龄 from dual "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取年龄"
    If Not rsTemp.EOF Then
        Get年龄 = Int(Nvl(rsTemp!年龄, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function

Public Function strProcess(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    strProcess = IIf(IsNull(varValue), DefaultValue, varValue)
    strProcess = Replace(strProcess, "'", "‘")
End Function

Private Function Save保险帐户_帐户余额(ByVal lng病人ID As Long, ByVal str动态信息 As String, ByRef dbl帐户余额 As Double) As Boolean
    '根据动态信息,算出帐户余额,并保存在保险帐户中
    '    写动动态信息：2005|252.07|370.20|20050318|39|0|39|0|0|0|0.00|2.5|0.00|0.00|0.00|0.00
    '    分解参数，取动态信息第2位，第3位，第7位，即：“|上年帐户结转金额|本年注入累计||||帐户支出累计|||||||||”
    '    本次帐户余额 = 上年帐户结转金额+本年注入累计-帐户支出累计＝252。07＋370.20－39＝583。27
    Dim strArr As Variant
    Save保险帐户_帐户余额 = False
    
    DebugTool "进入Save保险帐户_帐户余额函数:str动态信息=" & str动态信息
    strArr = Split(str动态信息 & "||||||||||", "|")
    dbl帐户余额 = Val(strArr(1)) + Val(strArr(2)) - Val(strArr(6))
    
    DebugTool "保存帐户余额:帐户余额=" & dbl帐户余额
    '更新保险帐户
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_黔南 & ",'帐户余额','" & Format(dbl帐户余额, "####0.0000;-####0;0000;0;0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊次数")
    DebugTool "保存帐户余额成功!"
    Save保险帐户_帐户余额 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
