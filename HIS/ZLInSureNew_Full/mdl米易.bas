Attribute VB_Name = "mdl米易"
Option Explicit
Private objCom As Object
Public gintComPort As Integer
Private mblnCreated As Boolean
Public Enum 字段
    个人编号 = 1
    卡号
    服务机构编号
    密码
    新密码
    支付金额
    经办人
    经办时间
    就诊编号
    结算编号
    支付类别
    人员类别
    记账流水号
    医保编码
    项目编码
    项目名称
    数量
    费用总额
    开单科室
    开单医生
    受单科室
    受单医生
    圈存机编号
    圈存金额
    圈存时间
    圈存流水号
    病种编码
    待遇获取时间
    入院日期
    入院诊断
    入院科室
    入院经办人
    入院经办时间
    结算时间
    跨年度结算标志
    真实结算标志
    出院原因
    出院日期
    出院诊断
    出院科室
    出院经办人
    出院经办时间
    退单编号
    就诊编号_起始
    就诊编号_截止
    记账流水号_起始
    记账流水号_截止
    结算编号_起始
    结算编号_截止
End Enum

Public gstrPara_米易 As String           '调用接口的参数串
Private Type ComInfo_米易
    系统时间 As String
    个人编号 As String
    卡号 As String
    服务机构编号 As String
    密码  As String
    新密码 As String
    支付金额 As Double
    就诊编号 As String
    结算编号 As String
    支付类别 As String
    圈存机编号 As String
    可圈存金额 As Double
    圈存金额 As Double
    圈存流水号 As String
    病种编码 As String
    待遇获取时间 As String
    入院日期 As String
    入院诊断 As String
    入院科室 As String
    记账流水号 As String
    出院原因 As String
    出院日期 As String
    出院诊断 As String
    出院科室 As String
    经办人 As String
    错误描述 As String
    错误代码 As Long
    交互错误信息 As String
    执行结果 As Long                '0表示发生错误；1表示正常完成
    姓名 As String
    性别 As String
    出生日期 As String
    身份证号 As String
    人员类别 As String
    单位名称 As String
    年龄 As Long
    帐户余额 As Double
    本次起付线 As Double
    本次统筹限额 As Double
    本次统筹报销比例 As Double
    挂钩自付 As Double
    进入统筹 As Double
    基数自付 As Double
    统筹自付 As Double
    统筹支付 As Double
    超限自付 As Double
    报销比例 As Double
'以下变量用于核对
    核对_记录数 As Double
    核对_帐户支付总额 As Double
    核对_医疗费总额 As Double
    核对_全自费总额 As Double
    核对_挂钩自费总额 As Double
    核对_进入统筹总额 As Double
    核对_现金支付总额 As Double
    核对_个人自付总额  As Double
    核对_在院人数 As Long
    核对_出院人数 As Long
    核对_数量 As Long
    核对_单价 As Double
    核对_自付比例 As Double
    核对_起付线支付金额 As Double
    核对_统筹自付金额 As Double
    核对_统筹支付金额 As Double
End Type
Public gComInfo_米易 As ComInfo_米易      '保存当前操作的数据
Private Const gint圈存流水号 As Integer = 1
Private Const gint记帐流水号 As Integer = 2
Private Const gint结算编号 As Integer = 3
'必须定义为LONG
Private Const glng圈存 As Long = 1
Private Const glng下帐 As Long = 0

Public Declare Function Card_Sale Lib "jpmyyy.dll" Alias "card_sale" _
(ByVal comport As Integer, ByVal userpsd As String, ByVal jystr As String, ByVal jymode As Integer) As Integer
'comport as 串口号(1-com1,2-com2,3-com3...)
'userpsd as 参保户使用密码，6位长度字符串,
'jystr as  41个字节长度 as jyje as 15位前台就诊记录号,2位操作员编码, 8位药店/医院编码,8位交易金额,8位刷卡金额,
'返回0 as 正确,
'功能 as 圈存、下帐接口程序,
Public Declare Function Card_ChangePsd Lib "jpmyyy.dll" Alias "change_psd" _
(ByVal comport As Integer, ByVal oldpsd As String, ByVal newpsd As String) As Integer
'comport as 串口号,oldpsd as 旧用户口令,newpsd as 修改后口令,
'返回值 as 0 as 修改正确，4 as  写卡错误, 3 as 转换数据错误,
' 2 as  读卡错误,1 as 输入原来口令不正确,
Public Declare Function Card_userinfo Lib "jpmyyy.dll" Alias "re_userinfo" _
(ByVal comport As Integer, ByVal userpsd As String, recode As Integer) As String
'卡片安全认证、返回10卡号，15个人顺序号，8卡上现金额(以分为单位)
'输入值，comport as 读写器串口号，userpsd as 使用密码，
'recode as 返回值 as 0成功，其他失败
'返回串 as recode为0返回的值正确，其它是汉字字符串错误信息,

'__________________________________________________________________________________
'需要保存：从卡内读出的社保号（与个人编号、卡号不同）及待遇审批时起付线，报销比例等
'需要在身份验证时: 显示待遇审批信息结算
'结算时，经办人（aae011）固定为"900090009000999"
'住院不支持中途结算 , 通过模拟结算实现
'必须以门诊的方式下个人帐户

Public Function 医保初始化_米易() As Boolean
    Dim rs服务机构编号  As New ADODB.Recordset
    Dim rs保险参数  As New ADODB.Recordset
    '功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
    '返回：初始化成功，返回true；否则，返回false
    
    On Error Resume Next
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
    Set rs服务机构编号 = zlDatabase.OpenSQLRecord(gstrSQL, "获取医院编码", type_米易)
    
    '由于圈存机编号与服务机构编号一致，所以同时对它赋值
    gComInfo_米易.服务机构编号 = Nvl(rs服务机构编号!医院编码, "")
    gComInfo_米易.圈存机编号 = gComInfo_米易.服务机构编号
    
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=[1]"
    Set rs保险参数 = zlDatabase.OpenSQLRecord(gstrSQL, "获取参数", type_米易)
    gintComPort = 1
    If Not rs保险参数.EOF Then gintComPort = Nvl(rs保险参数!参数值, 1)
    
    医保初始化_米易 = True
End Function

Public Function 医保终止_米易() As Boolean
    Debug.Print -1
    Set objCom = Nothing
    mblnCreated = False
    医保终止_米易 = True
End Function

Public Function 启动() As Boolean
    If mblnCreated Then Call 医保终止_米易
    
    mblnCreated = 创建对象
    If Not mblnCreated Then
        MsgBox "无法创建COM+对象，医保初始化失败！", vbInformation, gstrSysName
        Exit Function
    End If
    启动 = True
End Function

Public Function 医保设置_米易() As Boolean
    医保设置_米易 = frmSet米易.ShowME(type_米易)
End Function

Public Function 身份标识_米易(Optional bytType As Byte, Optional lng病人ID As Long) As String
    Dim str备注 As String, RSPATIENT As New ADODB.Recordset
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    If Not 启动 Then Exit Function
    身份标识_米易 = frmIdentify米易.GetPatient(bytType, lng病人ID)
    Call 医保终止_米易
    If 身份标识_米易 = "" Then Exit Function
    
    '强制把报销比例等数据填入
    gstrSQL = "Select 病人ID From 保险帐户 Where 险类=[1] And 医保号=[2]"
    Set RSPATIENT = zlDatabase.OpenSQLRecord(gstrSQL, "读取医保病人的基本信息", type_米易, gComInfo_米易.个人编号)
    lng病人ID = RSPATIENT!病人ID
    
    str备注 = Val(gComInfo_米易.本次起付线) & ";" & Val(gComInfo_米易.本次统筹报销比例) & ";" & Val(gComInfo_米易.本次统筹限额)
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & type_米易 & ",'备注','''" & str备注 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_米易")
End Function

Public Function 个人余额_米易(Optional ByVal bln读卡 As Boolean = False, Optional ByVal strSelfNo As String = "", Optional ByVal bln住院 As Boolean = False) As Currency
    '功能: 直接读出卡内金额
    '参数: 是否读卡
    '返回: 返回个人帐户余额
    Dim lng病人ID As Long
    Dim rsAcc As New ADODB.Recordset
    '读卡失败则退出
    gstrSQL = "Select Nvl(帐户余额,0) 帐户余额 From 保险帐户 Where 险类=[1]"
    If bln读卡 Then
        lng病人ID = ReadICCard(bln住院)
        If lng病人ID = 0 Then Exit Function
        '直接返回
        个人余额_米易 = gComInfo_米易.帐户余额 + gComInfo_米易.可圈存金额
        Exit Function
        
        gstrSQL = gstrSQL & " And 病人ID=[2]"
    Else
        gstrSQL = gstrSQL & " And 医保号=[3]"
    End If
    Set rsAcc = zlDatabase.OpenSQLRecord(gstrSQL, "读取帐户余额", type_米易, lng病人ID, strSelfNo)
    
    gComInfo_米易.帐户余额 = rsAcc!帐户余额
    个人余额_米易 = gComInfo_米易.帐户余额 + gComInfo_米易.可圈存金额
End Function

Private Function ReadICCard(Optional ByVal bln住院 As Boolean = False) As Long
    '读取卡内信息，同时更新结构体中病人相关信息、帐户余额，返回病人ID
    Dim recode As Integer, strResult As String, str就诊编号 As String
    Dim lng病人ID As Long
    Dim rsTemp As New ADODB.Recordset
    strResult = Card_userinfo(gintComPort, gComInfo_米易.密码, recode)
    If recode <> 0 Then
        MsgBox strResult, vbInformation, gstrSysName
        Exit Function
    End If
    
    gComInfo_米易.个人编号 = Mid(strResult, 11, 15)
    gComInfo_米易.卡号 = Mid(strResult, 1, 10)
    gComInfo_米易.帐户余额 = Val(Mid(strResult, 26, 8)) / 100                   '将分为单位记录的金额转换为以元为单位的金额
    
    str就诊编号 = gComInfo_米易.就诊编号
    '身份验证
    gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(卡号, gComInfo_米易.卡号) & _
        GetParaCode(服务机构编号, gComInfo_米易.服务机构编号)
    If Not 调用接口_米易("identifyinfogetting") Then
        Exit Function
    End If
    gComInfo_米易.就诊编号 = str就诊编号
    
    gstrSQL = "Select 病人ID From 保险帐户 Where 险类=[1] And 医保号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该医保病人的病人ID", type_米易, gComInfo_米易.个人编号)
    If rsTemp.EOF Then Exit Function
    ReadICCard = rsTemp!病人ID
End Function

Public Function WriteICCard(ByVal lng病人ID As Long, ByVal curMoney As Currency, Optional ByVal bln住院 As Boolean = False) As Boolean
    Dim blnRead As Boolean, blnErr As Boolean
    Dim lngReturn As Long
    Dim StrInput As String
    
    On Error GoTo errHand
    '写卡（上帐或下帐）
    If curMoney = 0 Then Exit Function
    
    '调用读卡
    blnRead = True
    Do While blnRead
        'ReadICCard:重新读取卡内余额及可圈存金额
        If lng病人ID <> ReadICCard(bln住院) Then
            MsgBox "读卡器内的卡不是当前病人的，请插入正确的卡后，按回车键！", vbInformation, gstrSysName
        Else
            blnRead = False
        End If
    Loop
    
    '要更新保险帐户，且同时更新结构体内的值
    '15位前台就诊记录号,2位操作员编码, 8位药店/医院编码,8位交易金额,8位刷卡金额
    StrInput = Abs(curMoney) * 100 '转换为分的格式
    If Len(StrInput) < 8 Then StrInput = String(8 - Len(StrInput), "0") & StrInput
    StrInput = Right(gComInfo_米易.就诊编号, 15) & "01" & gComInfo_米易.服务机构编号 & StrInput & StrInput
    If curMoney > 0 Then
        '下帐
        lngReturn = Card_Sale(gintComPort, gComInfo_米易.密码, StrInput, glng圈存)
    Else
        '上帐
        lngReturn = Card_Sale(gintComPort, gComInfo_米易.密码, StrInput, glng下帐)
    End If
    blnErr = (lngReturn = 0)
    gComInfo_米易.帐户余额 = gComInfo_米易.帐户余额 + curMoney
'    gComInfo_米易.圈存金额 = 0         '不能清为零，结算记录才能反应出卡内金额
    
    WriteICCard = blnErr
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 门诊虚拟结算_米易(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    Dim curTotal As Currency, cur个人帐户 As Currency
    Dim rsTemp As New ADODB.Recordset
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '注意：接口规定，门诊明细需结算后上传；住院明细需预结算时上传
    
    '身份验证后，返回的是卡外可圈存金额
    With rs明细
        '取出本次发生费用的金额合计
        Do While Not .EOF
            '先判断是否都设置了医保对应项目编码
            gstrSQL = " Select 项目编码 From 保险支付项目" & _
                      " Where 险类=[1] And 收费细目ID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否设置了对应的医保编码", type_米易, CLng(!收费细目ID))
            If rsTemp.EOF = True Then
                MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
                Exit Function
            End If
            
            curTotal = curTotal + CCur(Format(!实收金额, "#####0.00;-#####0.00;0;"))
            .MoveNext
        Loop
        
        gComInfo_米易.支付金额 = curTotal            '暂存费用总额
        If curTotal > gComInfo_米易.帐户余额 + gComInfo_米易.可圈存金额 Then
            cur个人帐户 = gComInfo_米易.帐户余额 + gComInfo_米易.可圈存金额
        Else
            cur个人帐户 = curTotal
        End If
        str结算方式 = "个人帐户;" & cur个人帐户 & ";1"   '允许修改
    End With
    
    '检查是否存在本次就诊编号，存在则提示不能结算，必须退出重新刷卡
    Dim strDate As String
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    gstrSQL = "Select count(*) Records From 保险结算记录 Where 性质=1 And 险类=[1] And 结算时间 Between [2] And [3] And 支付顺序号=[4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在本次就诊编号", type_米易, CDate(strDate), CDate(strDate & " 23:59:59"), gComInfo_米易.就诊编号)
    If rsTemp!Records <> 0 Then Exit Function
    
    门诊虚拟结算_米易 = True
End Function

Public Function 门诊结算_米易(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    Dim curTotal As Currency
    Dim int上传 As Integer
    Dim lng病人ID As Long
    Dim bln特种病 As Boolean, blnError As Boolean
    Dim rsTemp As New ADODB.Recordset
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '注意：接口规定，门诊明细需结算后上传；住院明细需预结算时上传，如果卡内金额不足，可以使用圈存接口，即将卡外的钱，调到卡内，以增加卡内金额
    '卡内余额需要通过卡操作函数读取，可圈存金额是接口返回，需要修改
    On Error GoTo errHand
    If Not 启动 Then Exit Function
    If Not 调用接口_米易("getsysdate") Then Exit Function
    
    Call 个人余额_米易(True, strSelfNo)
    
    '支付类别与医保病种是否是特种病有关
    bln特种病 = False
    gstrSQL = " Select A.病人ID,Nvl(B.类别,0) 特病种 " & _
              " From 保险帐户 A,(Select * From 保险病种 Where 险类=" & type_米易 & ") B " & _
              " Where A.险类=[1] And A.医保号=[2] And A.病种ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否是特种病", type_米易, strSelfNo)
    lng病人ID = rsTemp!病人ID
    bln特种病 = (rsTemp!特病种 <> 0)
    '门诊特种病不由医院端支付
    gComInfo_米易.支付类别 = "0201" 'IIf(bln特种病, "0205", "0201")
    gComInfo_米易.结算编号 = GetSequence(gint结算编号)
    gComInfo_米易.记账流水号 = GetSequence(gint记帐流水号)
    
    '调用圈存接口
    If cur个人帐户 > gComInfo_米易.帐户余额 Then
        gComInfo_米易.圈存金额 = cur个人帐户 - gComInfo_米易.帐户余额
        gComInfo_米易.圈存流水号 = Right(gComInfo_米易.服务机构编号, 3) & GetSequence(gint圈存流水号)
        
        gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(卡号, gComInfo_米易.卡号) & _
            GetParaCode(圈存机编号, gComInfo_米易.圈存机编号) & GetParaCode(圈存金额, gComInfo_米易.圈存金额) & _
            GetParaCode(经办时间, gComInfo_米易.系统时间) & GetParaCode(圈存流水号, gComInfo_米易.圈存流水号)
        If Not 调用接口_米易("qc") Then Exit Function
        If Not WriteICCard(lng病人ID, gComInfo_米易.圈存金额) Then
        '重复调用撤销接口，直到成功为止
            Do While True
                If 调用接口_米易("qcrollback") Then Exit Do
            Loop
            Exit Function
        End If
    End If
    
    '填写结算记录
    '累计进入统筹=本次圈存金额，帐户累计增加=原卡内金额，帐户累计支出=本次帐户支付
    '顺序号保存就诊编号，备注中保存当次的结算编号
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & type_米易 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & gComInfo_米易.帐户余额 - gComInfo_米易.圈存金额 & "," & cur个人帐户 & "," & gComInfo_米易.圈存金额 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_米易.支付金额 & "," & gComInfo_米易.支付金额 - cur个人帐户 & "," & 0 & "," & 0 & "," & 0 & ",0," & _
        0 & "," & cur个人帐户 & ",'" & gComInfo_米易.就诊编号 & "',null,null,'" & gComInfo_米易.结算编号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊收费数据")
    
    '如果使用了个人帐户，需要将使用情况上传（由于该接口无撤销函数，必需最后上传）
    If cur个人帐户 <> 0 Then
        If Not WriteICCard(lng病人ID, cur个人帐户 * -1) Then Exit Function
        gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & _
            GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & GetParaCode(支付金额, cur个人帐户) & _
            GetParaCode(经办时间, gComInfo_米易.系统时间) & GetParaCode(就诊编号, gComInfo_米易.就诊编号) & _
            GetParaCode(结算编号, gComInfo_米易.结算编号) & GetParaCode(支付类别, gComInfo_米易.支付类别)
        If Not 调用接口_米易("dataupload") Then
            '退出前将卡内金额还原
            'Modified by zyb 2003-10-25
            Do While 1
                If WriteICCard(lng病人ID, cur个人帐户) Then
                    Exit Do
                Else
                    Err.Raise 9000, gstrSysName, "写卡不成功，请插卡！", vbInformation, gstrSysName
                End If
            Loop
            Exit Function
        End If
    End If
    
    '上传费用明细记录
    blnError = False
    gstrSQL = "Select Rownum 标识号,A.ID,A.病人ID,A.NO,A.序号,A.记录性质,A.记录状态,A.登记时间,A.开单人 as 医生," & _
            "   A.数次*A.付数 as 数量,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额," & _
            "   A.收费类别,B.编码 as 项目编码,B.名称 as 项目名称,D.项目编码 医保编码," & _
            "   C.名称 开单部门,E.名称 受单部门" & _
            " From (Select * From 门诊费用记录 Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0) A,收费细目 B,部门表 C,保险支付项目 D,部门表 E " & _
            " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID(+) And A.执行部门ID=E.ID(+) And A.收费细目ID=D.收费细目ID And D.险类=[1]" & _
            " Order by A.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取本次结帐费用明细", type_米易)
    With rsTemp
        Do While Not .EOF
            gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & _
                GetParaCode(就诊编号, gComInfo_米易.就诊编号) & GetParaCode(记账流水号, gComInfo_米易.记账流水号 & !标识号) & _
                GetParaCode(结算编号, gComInfo_米易.结算编号) & GetParaCode(支付类别, gComInfo_米易.支付类别) & _
                GetParaCode(医保编码, !医保编码) & GetParaCode(项目编码, !项目编码) & GetParaCode(项目名称, !项目名称) & _
                GetParaCode(数量, !数量) & GetParaCode(费用总额, !结帐金额) & GetParaCode(开单科室, Nvl(!开单部门, "")) & _
                GetParaCode(开单医生, Nvl(!医生, "")) & GetParaCode(受单科室, Nvl(!受单部门, "")) & GetParaCode(受单医生, "") & _
                GetParaCode(经办时间, gComInfo_米易.系统时间)
            
            If 调用接口_米易("recipeinfotran") Then
                int上传 = 1
            Else
                int上传 = 0
                blnError = True
            End If
            
            '为病人费用记录打上标记，以便随时上传
            'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
            gstrSQL = "ZL_病人费用记录_更新医保(" & rsTemp("ID") & ",NULL,NULL,NULL,NULL," & int上传 & ",'" & gComInfo_米易.记账流水号 & !标识号 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
            .MoveNext
        Loop
    End With

    Call 医保终止_米易
    If blnError Then
        Err.Raise 9000, gstrSysName, "部分费用明细未正确上传，请到保险帐户管理中重新上传！", vbInformation, gstrSysName
    End If
    门诊结算_米易 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_米易(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    门诊结算冲销_米易 = False
End Function

Public Function 入院登记_米易(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    Dim str入院经办时间 As String
    Dim rsTemp As New ADODB.Recordset
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    
    On Error GoTo errHand
    If Not 启动 Then Exit Function
    
    gstrSQL = "Select A.登记人 经办人,B.名称 入院科室,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 入院经办时间," & _
            " to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 入院日期" & _
            " From 病案主页 A,部门表 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院信息", lng病人ID, lng主页ID)
    gComInfo_米易.经办人 = rsTemp!经办人
    gComInfo_米易.入院科室 = rsTemp!入院科室
    gComInfo_米易.入院日期 = rsTemp!入院日期
    str入院经办时间 = rsTemp!入院经办时间
    gComInfo_米易.入院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, True, False)
    
    gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & _
        GetParaCode(卡号, gComInfo_米易.卡号) & GetParaCode(密码, gComInfo_米易.密码) & _
        GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & GetParaCode(就诊编号, gComInfo_米易.就诊编号) & _
        GetParaCode(支付类别, gComInfo_米易.支付类别) & GetParaCode(病种编码, gComInfo_米易.病种编码) & _
        GetParaCode(入院日期, gComInfo_米易.入院日期) & GetParaCode(入院诊断, gComInfo_米易.入院诊断) & _
        GetParaCode(入院科室, gComInfo_米易.入院科室) & GetParaCode(入院经办人, gComInfo_米易.经办人) & _
        GetParaCode(入院经办时间, str入院经办时间)
    If Not 调用接口_米易("enterhospital") Then Exit Function
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & type_米易 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    Call 医保终止_米易
    入院登记_米易 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_米易(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim str出院经办时间 As String
    Dim bln医保出院 As Boolean
    Dim rsTemp As New ADODB.Recordset
    '功能：将出院信息发送医保前置服务器确认
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
    If Not 启动 Then Exit Function
    
    bln医保出院 = False
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        '调用医保的出院接口
        bln医保出院 = True
        gComInfo_米易.支付类别 = "0301"
        gComInfo_米易.出院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, False, False)
        '取出院原因
        gstrSQL = "select decode(出院方式,'正常',1,'转院',2,'死亡',3,9) 出院方式 From 病案主页 " & _
                " Where 病人ID = [1] And 主页ID = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院方式", lng病人ID, lng主页ID)
        gComInfo_米易.出院原因 = rsTemp!出院方式

        gstrSQL = "select b.名称 出院科室,床号,终止时间,操作员姓名  " & _
                 " from 病人变动记录 A,部门表 B  " & _
                 " where 病人ID=[1] and 终止原因=1 " & _
                 " and A.科室ID=B.ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院情况", lng病人ID)
        str出院经办时间 = Format(rsTemp!终止时间, "yyyy-MM-dd HH:mm:ss")
        gComInfo_米易.出院日期 = Format(rsTemp!终止时间, "yyyy-MM-dd")
        gComInfo_米易.出院科室 = ToVarchar(rsTemp!出院科室, 20)
        gComInfo_米易.经办人 = ToVarchar(rsTemp!操作员姓名, 20)
        gComInfo_米易.出院诊断 = ToVarchar(获取入出院诊断(lng病人ID, lng主页ID, False, False), 100)
        
        gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(卡号, gComInfo_米易.卡号) & _
                GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & GetParaCode(就诊编号, gComInfo_米易.就诊编号) & _
                GetParaCode(支付类别, gComInfo_米易.支付类别) & GetParaCode(出院原因, gComInfo_米易.出院原因) & _
                GetParaCode(出院日期, gComInfo_米易.出院日期) & GetParaCode(出院诊断, gComInfo_米易.出院诊断) & _
                GetParaCode(出院科室, gComInfo_米易.出院科室) & GetParaCode(出院经办人, gComInfo_米易.经办人) & _
                GetParaCode(出院经办时间, str出院经办时间)
        If Not 调用接口_米易("leavehospital") Then Exit Function
    End If
    
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & type_米易 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    
    Call 医保终止_米易
    MsgBox IIf(bln医保出院, "医保出院办理成功！", "HIS出院办理成功！"), vbInformation, gstrSysName
    出院登记_米易 = True
End Function

Public Function 入院登记撤销_米易(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
    If Not 启动 Then Exit Function
    gstrSQL = " Select Count(*) Records From 住院费用记录 " & _
              " Where 病人ID=[1] And 主页ID=[2] And Nvl(记录状态,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "撤销入院检查", lng病人ID, lng主页ID)
    If rsTemp!Records <> 0 Then
        MsgBox "已经存在费用记录，不允许办理撤销入院登记！", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & _
        GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & _
        GetParaCode(就诊编号, gComInfo_米易.就诊编号)
    If Not 调用接口_米易("enterhospitalrollback") Then Exit Function
    
    Call 医保终止_米易
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & type_米易 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    入院登记撤销_米易 = True
End Function

Public Function 出院登记撤销_米易(lng病人ID As Long, lng主页ID As Long) As Boolean
    '存在未结费用的病人才允许撤销HIS出院；否则认为已办理医保出院，不允许再办理HIS出院
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        MsgBox "医保已出院的病人不允许撤销出院！", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & type_米易 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销出院登记")
    出院登记撤销_米易 = True
End Function

Public Function 住院虚拟结算_米易(rsExse As Recordset, ByVal lng病人ID As Long) As String
    Dim curTotal As Currency
    Dim lng主页ID As Long
    Dim cur个人自付 As Currency, cur个人帐户 As Currency
    Dim str入院年份 As String, str结算年份 As String
    Dim str结算时间 As String, str经办时间 As String
    Dim blnUpload As Boolean
    Dim rsTemp As New ADODB.Recordset
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    '接口返回的报销额减去本次住院期间以往报销额的汇总金额后，才是本次的实际报销额
    'rsExse记录集中的字段清单
    '记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
    On Error GoTo errHand
    If Not 启动 Then Exit Function
    
    If Not 调用接口_米易("getsysdate") Then Exit Function
    Call 获取病人相关信息(lng病人ID)
    Call 个人余额_米易(True, "", True)
    cur个人帐户 = gComInfo_米易.帐户余额
    
    gstrSQL = " Select B.住院次数 主页ID,to_char(A.入院日期,'yyyy') 入院年份 " & _
              " From 病案主页 A,病人信息 B" & _
              " Where B.病人ID=[1] And A.主页ID=B.住院次数 And A.病人ID=B.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人入院时间", lng病人ID)
    str入院年份 = rsTemp!入院年份
    lng主页ID = rsTemp!主页ID
    
    gComInfo_米易.结算编号 = 0               '上传费用明细时，结算编号必需要置为零
    gComInfo_米易.支付类别 = "0301"
    gComInfo_米易.记账流水号 = GetSequence(gint记帐流水号)
    str经办时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str结算时间 = Format(gComInfo_米易.系统时间, "yyyy-MM-dd HH:mm:ss")
    str结算年份 = Format(gComInfo_米易.系统时间, "yyyy")
    
    With rsExse
        Do While Not .EOF
            If IsNull(!医保项目编码) Then
                MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
                Exit Function
            End If
            .MoveNext
        Loop
        .MoveFirst
        
        '上传费用明细
        curTotal = 0
        Do While Not .EOF
            curTotal = curTotal + !金额
            
            blnUpload = True
            If Not IsNull(!是否上传) Then
                blnUpload = (!是否上传 = 0)
            End If
            
            If blnUpload Then
                
                '取该收费细目的编码与名称
                gstrSQL = "Select 编码 项目编码 ,名称 项目名称 From 收费细目 Where ID = [1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该收费细目的编码与名称", CLng(!收费细目ID))
                
                gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & _
                    GetParaCode(就诊编号, gComInfo_米易.就诊编号) & GetParaCode(记账流水号, gComInfo_米易.记账流水号 & .AbsolutePosition) & _
                    GetParaCode(结算编号, gComInfo_米易.结算编号) & GetParaCode(支付类别, gComInfo_米易.支付类别) & _
                    GetParaCode(医保编码, !医保项目编码) & GetParaCode(项目编码, rsTemp!项目编码) & GetParaCode(项目名称, rsTemp!项目名称) & _
                    GetParaCode(数量, !数量) & GetParaCode(费用总额, !金额) & GetParaCode(开单科室, Nvl(!开单部门, "")) & _
                    GetParaCode(开单医生, Nvl(!医生, "")) & GetParaCode(受单科室, "") & GetParaCode(受单医生, "") & _
                    GetParaCode(经办时间, Format(!发生时间, "yyyy-MM-dd HH:mm:ss"))
                If 调用接口_米易("recipeinfotran") Then
                    '仅打上传标志，因为明细必须正确上传后，才能保证结算的正确性
                    'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                    gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
                Else
                    Exit Function
                End If
            End If
            .MoveNext
        Loop
        
        gComInfo_米易.支付金额 = curTotal
        '虚拟结算（返回的数据减去历次结算数据，就等于本次的真实结算数据）
        gComInfo_米易.结算编号 = GetSequence(gint结算编号)
        gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(卡号, gComInfo_米易.卡号) & _
            GetParaCode(密码, gComInfo_米易.密码) & GetParaCode(就诊编号, gComInfo_米易.就诊编号) & _
            GetParaCode(结算编号, gComInfo_米易.结算编号) & GetParaCode(费用总额, gComInfo_米易.支付金额) & _
            GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & GetParaCode(支付类别, gComInfo_米易.支付类别) & _
            GetParaCode(病种编码, gComInfo_米易.病种编码) & GetParaCode(经办人, "900090009000999") & _
            GetParaCode(结算时间, str结算时间) & GetParaCode(经办时间, str经办时间) & _
            GetParaCode(跨年度结算标志, IIf(str入院年份 <> str结算年份, 1, 0)) & GetParaCode(真实结算标志, 0)
        If Not 调用接口_米易("ExpenseReckoning") Then Exit Function
        
        Call 费用分隔(lng病人ID, lng主页ID)
    End With
    
    '返回结算方式
    cur个人自付 = gComInfo_米易.支付金额 - gComInfo_米易.统筹支付 'gComInfo_米易.统筹自付 + gComInfo_米易.基数自付 + gComInfo_米易.超限自付 + gComInfo_米易.挂钩自付
    If gComInfo_米易.统筹支付 <> 0 Then
        住院虚拟结算_米易 = "医保基金;" & gComInfo_米易.统筹支付 & ";0"
    End If
    '只有出院结算才允许使用个人帐户
    If 医保病人已经出院(lng病人ID) Then
        If cur个人帐户 <> 0 Then
            If cur个人帐户 > cur个人自付 Then
                cur个人帐户 = cur个人自付
            End If
            If cur个人帐户 < 0 Then cur个人帐户 = 0
            住院虚拟结算_米易 = 住院虚拟结算_米易 & IIf(住院虚拟结算_米易 = "", "", "|") & "个人帐户;" & cur个人帐户 & ";1"
        End If
    End If
    
    Call 医保终止_米易
    If 住院虚拟结算_米易 = "" Then 住院虚拟结算_米易 = "个人帐户;0;1"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_米易(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    Dim cur个人帐户 As Currency
    Dim lng主页ID As Long
    Dim blnError As Boolean
    Dim str入院年份 As String, str结算年份 As String
    Dim str经办时间 As String, str结算时间 As String
    Dim str就诊编号 As String
    Dim rsTemp As New ADODB.Recordset
    '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
        '虚拟结算（返回的数据减去历次结算数据，就等于本次的真实结算数据）
    On Error GoTo errHand
    If Not 启动 Then Exit Function
    If Not 调用接口_米易("getsysdate") Then Exit Function
    Call 个人余额_米易(True, "", True)
    cur个人帐户 = gComInfo_米易.帐户余额
    
    gstrSQL = " Select B.住院次数 主页ID,to_char(A.入院日期,'yyyy') 入院年份 " & _
              " From 病案主页 A,病人信息 B" & _
              " Where B.病人ID=[1] And A.主页ID=B.住院次数 And A.病人ID=B.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人入院时间", lng病人ID)
    str入院年份 = rsTemp!入院年份
    lng主页ID = rsTemp!主页ID
    
    str经办时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str结算时间 = Format(gComInfo_米易.系统时间, "yyyy-MM-dd HH:mm:ss")
    str结算年份 = Format(gComInfo_米易.系统时间, "yyyy")
    
    '住院结算
    gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(卡号, gComInfo_米易.卡号) & _
        GetParaCode(密码, gComInfo_米易.密码) & GetParaCode(就诊编号, gComInfo_米易.就诊编号) & _
        GetParaCode(结算编号, gComInfo_米易.结算编号) & GetParaCode(费用总额, gComInfo_米易.支付金额) & _
        GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & GetParaCode(支付类别, gComInfo_米易.支付类别) & _
        GetParaCode(病种编码, gComInfo_米易.病种编码) & GetParaCode(经办人, 1E+21) & _
        GetParaCode(结算时间, str结算时间) & GetParaCode(经办时间, str经办时间) & _
        GetParaCode(跨年度结算标志, IIf(str入院年份 <> str结算年份, 1, 0)) & GetParaCode(真实结算标志, 1)
    If Not 调用接口_米易("ExpenseReckoning") Then Exit Function
    Call 费用分隔(lng病人ID, lng主页ID)
    
    '读取本次个人帐户支付额
    gstrSQL = "Select Nvl(A.冲预交,0) 个人帐户 " & _
        " From 病人预交记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And B.险类=" & type_米易 & _
        " And A.结算方式 in ('个人帐户') And A.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取本次个人帐户支付额", lng结帐ID)
    cur个人帐户 = 0
    If Not rsTemp.EOF Then
        cur个人帐户 = rsTemp!个人帐户
    End If
    
    '填写保险结算记录
    '累计进入统筹=本次圈存金额，帐户累计增加=原卡内金额，帐户累计支出=本次帐户支付
    '顺序号保存就诊编号，备注中保存当次的结算编号
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & type_米易 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & gComInfo_米易.帐户余额 & "," & cur个人帐户 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & gComInfo_米易.基数自付 & "," & _
        gComInfo_米易.支付金额 & "," & 0 & "," & gComInfo_米易.挂钩自付 & "," & gComInfo_米易.统筹支付 + gComInfo_米易.统筹自付 & "," & gComInfo_米易.统筹支付 & ",0," & _
        gComInfo_米易.超限自付 & "," & cur个人帐户 & ",'" & gComInfo_米易.就诊编号 & "',null,null,'" & gComInfo_米易.结算编号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算数据")
    
    '如果使用了个人帐户，需要将使用情况上传（由于该接口无撤销函数，必须放在最后上传）
    '使用个人帐户的地方，注意写卡
    If cur个人帐户 <> 0 Then
        '必需以门诊的方式下个人帐户
        blnError = False
        If Not WriteICCard(lng病人ID, cur个人帐户 * -1, True) Then blnError = True
        
        str就诊编号 = gComInfo_米易.就诊编号
        gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(卡号, gComInfo_米易.卡号) & _
            GetParaCode(服务机构编号, gComInfo_米易.服务机构编号)
        If 调用接口_米易("identifyinfogetting") Then
            gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & _
                GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & GetParaCode(支付金额, cur个人帐户) & _
                GetParaCode(经办时间, gComInfo_米易.系统时间) & GetParaCode(就诊编号, gComInfo_米易.就诊编号) & _
                GetParaCode(结算编号, gComInfo_米易.结算编号) & GetParaCode(支付类别, "0201")
            If blnError = False Then
                If Not 调用接口_米易("dataupload") Then
                    'Modified by zyb 2003-10-25
                    Do While 1
                        If WriteICCard(lng病人ID, cur个人帐户, True) Then
                            Exit Do
                        Else
                            Err.Raise 9000, gstrSysName, "写卡不成功，请插卡！", vbInformation, gstrSysName
                        End If
                    Loop
                    blnError = True
                End If
            End If
        Else
            blnError = True
        End If
        
        gComInfo_米易.就诊编号 = str就诊编号
        If blnError Then
            '不能退出，提示以现金支付
            gcnOracle.Execute "Delete 保险结算记录 Where 性质=2 And 险类=" & type_米易
            
            gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & type_米易 & "," & lng病人ID & "," & _
                Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
                0 & "," & "NULL" & "," & 0 & "," & 0 & "," & gComInfo_米易.基数自付 & "," & _
                gComInfo_米易.支付金额 & "," & 0 & "," & gComInfo_米易.挂钩自付 & "," & gComInfo_米易.统筹支付 + gComInfo_米易.统筹自付 & "," & gComInfo_米易.统筹支付 & ",0," & _
                gComInfo_米易.超限自付 & "," & 0 & ",'" & gComInfo_米易.就诊编号 & "',null,null,'" & gComInfo_米易.结算编号 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算数据")
            Err.Raise 9000, gstrSysName, "下个人帐户失败，注意收取现金（" & Format(cur个人帐户, "#####0.00;-#####0.00; ;") & "）", vbInformation, gstrSysName
        End If
    End If
    Call 医保终止_米易
    
    '只有出院结算才允许使用个人帐户
    If 医保病人已经出院(lng病人ID) Then
        Call 出院登记_米易(lng病人ID, lng主页ID)
    End If
    住院结算_米易 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_米易(lng结帐ID As Long) As Boolean
    Dim lng冲销ID As Long
    Dim str退单编号 As String
    Dim rsTemp As New ADODB.Recordset
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    On Error GoTo errHand
    If Not 启动 Then Exit Function
    
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "米易医保", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "Select * " & _
              "  From 保险结算记录 Where 性质=2 and 记录ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "米易医保", lng结帐ID)
    
    Call 获取病人相关信息(rsTemp!病人ID)
    gComInfo_米易.结算编号 = GetSequence(gint结算编号)   '取新的结算编号
    gComInfo_米易.就诊编号 = Nvl(rsTemp!顺序号, "")      '取当时的就诊编号
    str退单编号 = Nvl(rsTemp!备注, "")              '取当时的结算编号
    
    '填写结算记录
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & type_米易 & "," & rsTemp!病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & Nvl(rsTemp!帐户累计增加, 0) * -1 & "," & Nvl(rsTemp!帐户累计支出, 0) * -1 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & Nvl(rsTemp!实际起付线, 0) * -1 & "," & _
        Nvl(rsTemp!发生费用金额, 0) * -1 & "," & 0 & "," & Nvl(rsTemp!首先自付金额, 0) * -1 & "," & Nvl(rsTemp!进入统筹金额, 0) * -1 & "," & Nvl(rsTemp!统筹报销金额, 0) * -1 & ",0," & _
        Nvl(rsTemp!超限自付金额, 0) * -1 & "," & Nvl(rsTemp!个人帐户支付, 0) * -1 & ",'" & gComInfo_米易.就诊编号 & "',null,null,'" & gComInfo_米易.结算编号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算数据")
    
    '调用撤销结算接口
    gstrPara_米易 = GetParaCode(就诊编号, gComInfo_米易.就诊编号) & GetParaCode(结算编号, gComInfo_米易.结算编号) & _
            GetParaCode(退单编号, str退单编号) & GetParaCode(服务机构编号, gComInfo_米易.服务机构编号)
    If Not 调用接口_米易("expenserollback") Then Exit Function
    
    Call 医保终止_米易
    住院结算冲销_米易 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Private Sub 费用分隔(ByVal lng病人ID As Long, lng主页ID As Long)
    Dim cur费用总额 As Currency, cur统筹总额 As Currency, cur统筹自付总额 As Currency '历次中途结算费用总额;历次中途结算统筹支付总额
    Dim cur起付线总额 As Currency, cur挂钩自付总额 As Currency, cur超限自付总额 As Currency
    Dim rsTemp As New ADODB.Recordset
    
    '取本次住院期间，历次中途结算的费用总额及统筹总额
    gstrSQL = "SELECT SUM(发生费用金额) 发生费用金额,SUM(进入统筹金额) 进入统筹金额,SUM(统筹报销金额) 统筹报销金额, " & _
             " SUM(首先自付金额) 首先自付金额,SUM(实际起付线) 实际起付线,SUM(超限自付金额) 超限自付金额" & _
             " FROM  " & _
             "      (SELECT Distinct 病人ID,结帐ID FROM 住院费用记录 " & _
             "      WHERE 病人ID=[1] AND 主页ID= [2]" & _
             "      ) A,保险结算记录 B " & _
             " WHERE A.病人ID=B.病人ID AND B.记录ID=A.结帐ID AND B.险类=[3] AND B.性质=2 " & _
             " GROUP BY A.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取本次住院期间以往结算的费用总额及统筹报销总额", lng病人ID, lng主页ID, type_米易)
    If Not rsTemp.EOF Then
        cur费用总额 = rsTemp!发生费用金额
        cur统筹总额 = rsTemp!统筹报销金额
        cur统筹自付总额 = rsTemp!进入统筹金额 - rsTemp!统筹报销金额
        cur起付线总额 = rsTemp!实际起付线
        cur挂钩自付总额 = rsTemp!挂钩自付金额
        cur超限自付总额 = rsTemp!超限自付金额
    Else
        cur费用总额 = 0
        cur统筹总额 = 0
        cur统筹自付总额 = 0
        cur起付线总额 = 0
        cur挂钩自付总额 = 0
        cur超限自付总额 = 0
    End If
    
    gComInfo_米易.挂钩自付 = CCur(Format(gComInfo_米易.挂钩自付 - cur挂钩自付总额, "#####0.00;-#####0.00;0;"))
    gComInfo_米易.进入统筹 = CCur(Format(gComInfo_米易.进入统筹 - (cur统筹总额 + cur统筹自付总额), "#####0.00;-#####0.00;0;"))
    gComInfo_米易.基数自付 = CCur(Format(gComInfo_米易.基数自付 - cur起付线总额, "#####0.00;-#####0.00;0;"))
    gComInfo_米易.统筹支付 = CCur(Format(gComInfo_米易.统筹支付 - cur统筹总额, "#####0.00;-#####0.00;0;"))
    gComInfo_米易.统筹自付 = CCur(Format(gComInfo_米易.统筹自付 - cur统筹自付总额, "#####0.00;-#####0.00;0;"))
    gComInfo_米易.超限自付 = CCur(Format(gComInfo_米易.超限自付 - cur超限自付总额, "#####0.00;-#####0.00;0;"))
End Sub

Public Function 调用接口_米易(ByVal strFunction As String) As Boolean
    '调用接口功能
    On Error GoTo errHand
    
    Select Case strFunction
    Case "getsysdate"                   '获取系统时间
        Dim strSysdate As Date
        Call objCom.getsysdate(strSysdate, gComInfo_米易.执行结果)
        gComInfo_米易.系统时间 = Format(strSysdate, "yyyy-MM-dd HH:mm:ss")
    Case "identifyinfogetting"          '身份验证
        Call objCom.identifyinfogetting(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, _
        gComInfo_米易.执行结果, gComInfo_米易.就诊编号, gComInfo_米易.姓名, gComInfo_米易.性别, strSysdate, gComInfo_米易.身份证号, _
        gComInfo_米易.人员类别, gComInfo_米易.单位名称, gComInfo_米易.年龄, gComInfo_米易.可圈存金额)
        gComInfo_米易.出生日期 = Format(strSysdate, "yyyy-MM-dd")
    Case "modifypassword"               '修改密码
        Call objCom.modifypassword(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果)
    Case "dataupload"                   '上传结算数据（帐户）
        Call objCom.dataupload(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果)
    Case "recipeinfotran"               '上传费用明细
        Call objCom.recipeinfotran(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果)
    Case "qc"                           'IC卡圈存
        Call objCom.qc(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果)
    Case "qcrollback"                   'IC卡圈存撤销
        Call objCom.qcrollback(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果)
    Case "audittreatment"               '资格审批待遇核定
        Call objCom.audittreatment(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果, _
        gComInfo_米易.就诊编号, gComInfo_米易.姓名, gComInfo_米易.性别, strSysdate, gComInfo_米易.身份证号, gComInfo_米易.人员类别, _
        gComInfo_米易.年龄, gComInfo_米易.本次起付线, gComInfo_米易.本次统筹限额, gComInfo_米易.本次统筹报销比例, gComInfo_米易.单位名称)
        gComInfo_米易.出生日期 = Format(strSysdate, "yyyy-MM-dd")
    Case "enterhospital"                '入院办理
        Call objCom.enterhospital(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果)
    Case "enterhospitalrollback"        '入院办理撤销
        Call objCom.enterhospitalrollback(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果)
    Case "ExpenseReckoning"             '住院结算/虚拟结算
        Call objCom.ExpenseReckoning(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果, _
        gComInfo_米易.挂钩自付, gComInfo_米易.进入统筹, gComInfo_米易.基数自付, gComInfo_米易.统筹支付, gComInfo_米易.统筹自付, gComInfo_米易.超限自付, gComInfo_米易.报销比例)
    Case "leavehospital"                '出院办理
        Call objCom.leavehospital(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果)
    Case "expenserollback"              '住院结算撤销
        Call objCom.expenserollback(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果)
    Case "checkaccount"                 '核对个人帐户支付
        Call objCom.checkaccount(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果, _
        gComInfo_米易.核对_记录数, gComInfo_米易.核对_帐户支付总额)
    Case "checkexpense"                 '核对门诊结算信息
        Call objCom.checkexpense(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果, _
        gComInfo_米易.核对_记录数, gComInfo_米易.核对_医疗费总额, gComInfo_米易.核对_全自费总额, gComInfo_米易.核对_挂钩自费总额, _
        gComInfo_米易.核对_进入统筹总额, gComInfo_米易.核对_帐户支付总额, gComInfo_米易.核对_现金支付总额, gComInfo_米易.核对_个人自付总额)
    Case "checkenterleavehosptinfo"     '核对住院人次
        Call objCom.checkenterleavehosptinfo(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果, _
        gComInfo_米易.核对_在院人数, gComInfo_米易.核对_出院人数)
    Case "checkrecipeinfo"              '核对费用明细
        Call objCom.checkrecipeinfo(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果, _
        gComInfo_米易.核对_数量, gComInfo_米易.核对_单价, gComInfo_米易.核对_医疗费总额, gComInfo_米易.核对_自付比例, gComInfo_米易.核对_全自费总额, _
        gComInfo_米易.核对_挂钩自费总额, gComInfo_米易.核对_进入统筹总额)
    Case "checkexpensereckoninginfo"    '核对住院费用结算结果
        Call objCom.checkexpensereckoninginfo(gstrPara_米易, gComInfo_米易.错误描述, gComInfo_米易.错误代码, gComInfo_米易.交互错误信息, gComInfo_米易.执行结果, _
        gComInfo_米易.核对_起付线支付金额, gComInfo_米易.核对_统筹自付金额, gComInfo_米易.核对_统筹支付金额)
    End Select
    
    If gComInfo_米易.执行结果 = 0 Then
        MsgBox gComInfo_米易.错误描述 & "|" & gComInfo_米易.交互错误信息 & "|错误代码：" & gComInfo_米易.错误代码, vbInformation, gstrSysName
        Exit Function
    End If
    
    调用接口_米易 = True
    Exit Function
errHand:
    MsgBox "交易执行失败！", vbInformation, gstrSysName
End Function

Public Function 创建对象() As Boolean
    On Error GoTo errHand
    
    Set objCom = CreateObject("pb80.n_center_interface.1.0")
    If objCom Is Nothing Then
        Exit Function
    End If
    
    创建对象 = True
    Exit Function
errHand:
End Function

Public Function 获取病人相关信息(ByVal lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '读取医保病人相关信息，并更新公用结构体
    gstrSQL = " Select 卡号,密码,医保号 个人编号,顺序号 就诊编号,Nvl(病种ID,0) 病种ID From 保险帐户" & _
            " Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医保病人的相关信息", type_米易, lng病人ID)
    If rsTemp.EOF Then Exit Function
    
    gComInfo_米易.个人编号 = rsTemp!个人编号
    gComInfo_米易.卡号 = rsTemp!卡号
    gComInfo_米易.就诊编号 = rsTemp!就诊编号
    gComInfo_米易.密码 = rsTemp!密码
    Call 获取病种编码(rsTemp!病种ID)
    获取病人相关信息 = True
End Function

Public Sub 获取病种编码(ByVal lng病种ID As Long)
    Dim rsTemp As New ADODB.Recordset
    '判断病种类别，如果是慢特病，则病种编码="900001"；否则="900002"
    gComInfo_米易.病种编码 = "900002"
    gstrSQL = "Select nvl(类别,0) 类别 From 保险病种 Where 险类=[1] And ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病种信息", type_米易, lng病种ID)
    If Not rsTemp.EOF Then
        If rsTemp!类别 <> 0 Then
            gComInfo_米易.病种编码 = "900001"
        End If
    End If
End Sub

Public Function GetSequence(Optional ByVal intType As Integer = 1) As String
    'intTYPE的含义：
    '1=圈存流水号=医院编号+7位流水号（用于圈存及圈存失败时使用，需要保存）
    '2=记帐流水号=15位数字流水号（可以不保存在数据库中）
    '3=结算编号=J+YYMMDD+医院编号+4位流水号（同一就诊编号不允许重复，需要保存）
    Dim strSequence As String
    Dim strHour As String, strMinute As String, strSecond As String
    Dim rsTemp As New ADODB.Recordset
    
    '按规则产生各自的流水号
'    gComInfo_米易.系统时间 = Now
    strHour = Format(gComInfo_米易.系统时间, "HH")
    strMinute = Mid(gComInfo_米易.系统时间, 15, 2) ' Format(gComInfo_米易.系统时间, "mm")
    strSecond = Format(gComInfo_米易.系统时间, "ss")
    Select Case intType
'    Case 1      '和单据流水号产生规则一致（A(年)+A(月)+A(日)+HHmm）
'        strSequence = 获取一位表示串(1) & 获取一位表示串(2) & 获取一位表示串(3) & 获取一位表示串(4, strHour) & 获取一位表示串(4, strMinute) & strSecond
    Case 1, 2     '以当前系统时间的yyMMddHHmmss，后三位以当前记录的序号字段的值填充
        strSequence = Format(gComInfo_米易.系统时间, "yyMMddHHmmss")
    Case 3      '4位流水号以当时的HHmm为标识
        strSequence = "J" & Format(gComInfo_米易.系统时间, "yyMMdd") & gComInfo_米易.服务机构编号 & 获取一位表示串(4, strHour) & 获取一位表示串(4, strMinute) & strSecond
    End Select
    GetSequence = strSequence
End Function

Public Function 获取一位表示串(Optional ByVal intType As Integer = 1, Optional ByVal lngData As Long = 0) As String
    Dim lngMid As Long
    '返回一位的年份、月份或日子的表示字符串；1-年份;2-月份;3-日子
    Select Case intType
    Case 1
        lngMid = Format(gComInfo_米易.系统时间, "yyyy")
        lngMid = lngMid - 2000
    Case 2
        lngMid = Format(gComInfo_米易.系统时间, "MM")
    Case 3
        lngMid = Format(gComInfo_米易.系统时间, "dd")
    Case Else
        lngMid = lngData
    End Select
    If lngMid >= 10 Then
        获取一位表示串 = Chr(lngMid - 10 + 65)
    Else
        获取一位表示串 = lngMid
    End If
End Function

Public Function GetParaCode(ByVal intType As Integer, ByVal strData As Variant) As String
'项目编码    参数值       项目类型        项目意义
'AKA123         0       特殊病种标志    非特殊病种
'AKA123         1       特殊病种标志    特殊病种
'AKC021         11      医疗人员类别    在职
'AKC021         12      医疗人员类别    在职长期驻外
'AKC021         21      医疗人员类别    退休
'AKC021         22      医疗人员类别    退休异地安置
'AKC021         31      医疗人员类别    离休
'AAC004         1       性别            男
'AAC004         2       性别            女
'AKA130         0101    支付类别        药店购药
'AKA130         0201    支付类别        普通门诊
'AKA130         0205    支付类别        特殊病种门诊
'AKA130         0301    支付类别        普通住院
'AKA130         0302    支付类别        非居住地住院
'AKA130         0304    支付类别        统筹区内转院
'AKA130         0305    支付类别        转外住院
'AKA130         0401    支付类别        跨年度住院(只能在ka10k1中使用)
'AKA130         0901    支付类别        定点医疗机构预付
'AKA130         0902    支付类别        定点医疗机构罚款
'AKA130         0701    支付类别        公务员报销类别
'YKA002         2000001 医疗项目编码    甲类
'YKA002         2000002 医疗项目编码    乙类
'YKA002         2000003 医疗项目编码    自费
'YKA002         2000004 医疗项目编码    一级医院床位费
'YKA002         2000005 医疗项目编码    二级医院床位费
'YKA002         2000006 医疗项目编码    三级医院床位费
'YKA026         900001  病种编码        医保规定住院病种（起付线＝0）
'YKA026         900002  病种编码        非规定住院病种（存在起付线）
'AKC195         1       出院原因        康复
'AKC195         2       出院原因        转院
'AKC195         3       出院原因        死亡
'AKC195         9       出院原因        其它
    
    Dim strValue As String
    Select Case intType
    Case 个人编号
        strValue = "aac001"
    Case 卡号
        strValue = "yac005"
    Case 服务机构编号
        strValue = "akb020"
    Case 密码
        strValue = "ykc005"
    Case 新密码
        strValue = "new_ykc005"
    Case 支付金额
        strValue = "defrayamount"
    Case 经办人
        strValue = "aae011"
    Case 经办时间
        strValue = "aae036"
    Case 就诊编号
        strValue = "akc190"
    Case 结算编号
        strValue = "yka103"
    Case 支付类别
        strValue = "aka130"
    Case 记账流水号
        strValue = "yka105"
    Case 医保编码
        strValue = "yka002"
    Case 项目编码
        strValue = "yka094"
    Case 项目名称
        strValue = "yka095"
    Case 数量
        strValue = "akc226"
    Case 费用总额
        strValue = "yka055"
    Case 开单科室
        strValue = "yka098"
    Case 开单医生
        strValue = "yka099"
    Case 受单科室
        strValue = "yka101"
    Case 受单医生
        strValue = "yka102"
    Case 圈存机编号
        strValue = "yka151"
    Case 圈存金额
        strValue = "yka152"
    Case 圈存时间
        strValue = "yka153"
    Case 圈存流水号
        strValue = "ykc019"
    Case 病种编码
        strValue = "yka026"
    Case 待遇获取时间, 结算时间, 出院日期
        strValue = "akc194"
    Case 入院日期
        strValue = "akc192"
    Case 入院诊断
        strValue = "akc193"
    Case 入院科室
        strValue = "ykc011"
    Case 入院经办人
        strValue = "ykc013"
    Case 入院经办时间
        strValue = "ykc014"
    Case 跨年度结算标志
        strValue = "ykc007"
    Case 真实结算标志
        strValue = "mnjs"
    Case 出院原因
        strValue = "akc195"
    Case 出院诊断
        strValue = "akc196"
    Case 出院科室
        strValue = "ykc015"
    Case 出院经办人
        strValue = "ykc017"
    Case 出院经办时间
        strValue = "ykc018"
    Case 退单编号
        strValue = "yka198"
    Case 就诊编号_起始
        strValue = "akc190_Begin"
    Case 就诊编号_截止
        strValue = "akc190_End"
    Case 记账流水号_起始
        strValue = "yka105_begin"
    Case 记账流水号_截止
        strValue = "yka105_end"
    Case 结算编号_起始
        strValue = "yka103_begin"
    Case 结算编号_截止
        strValue = "yka103_end"
    Case 人员类别
        strValue = "akc021"
    End Select
    
    GetParaCode = "<" & strValue & ">" & strData & "</" & strValue & ">"
End Function

Public Sub 核对帐户支付_米易()
    Dim cur帐户支付总额 As Currency
    Dim str开始日期 As String, str结束日期 As String
    Dim str开始就诊编号 As String, str结束就诊编号 As String
    Dim str开始结算编号 As String, str结束结算编号 As String
    Dim rsAccount As New ADODB.Recordset
    
    On Error GoTo errHand
    '获取查询日期
    If frm日期范围_米易.Show_ME(str开始日期, str结束日期) = False Then Exit Sub
    '本地提取帐户支付总额
    gstrSQL = "Select SUM(冲预交) 个人帐户 " & _
        " From 病人预交记录 " & _
        " Where 结算方式 in ('个人帐户') " & _
        " And 收款时间 Between [1] And [2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "统计帐户支付总额", CDate(str开始日期), CDate(str结束日期))
    If Not rsAccount.EOF Then
        cur帐户支付总额 = Nvl(rsAccount!个人帐户, 0)
    End If
    
    If Not 启动 Then Exit Sub
    '获取指定日期范围内的开始、结束就诊编号
    If Not FUNC_就诊编号(str开始日期, str结束日期, str开始就诊编号, str结束就诊编号, _
    str开始结算编号, str结束结算编号, False) Then Exit Sub
    '获取医保中心返回的帐户支付总额
    gstrPara_米易 = GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & _
        GetParaCode(就诊编号_起始, str开始就诊编号) & GetParaCode(就诊编号_截止, str结束就诊编号)
    Call 调用接口_米易("checkaccount")
    Call 医保终止_米易
    
    If Format(cur帐户支付总额, "#####0.00;-#####0.00;0;") <> Format(gComInfo_米易.核对_帐户支付总额, "#####0.00;-#####0.00;0;") Then
        MsgBox "（本地）帐户支付总额：" & cur帐户支付总额 & String(4, " ") & "（医保）帐户支付总额：" & gComInfo_米易.核对_帐户支付总额
    Else
        MsgBox "数据正确无误，核对成功！", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub 核对门诊结算_米易()
    Dim cur发生费用金额 As Currency, cur首先自付金额 As Currency, cur个人帐户支付 As Currency
    Dim str开始日期 As String, str结束日期 As String
    Dim str开始就诊编号 As String, str结束就诊编号 As String
    Dim str开始结算编号 As String, str结束结算编号 As String
    Dim rsAccount As New ADODB.Recordset
    
    On Error GoTo errHand
    '获取查询日期
    If frm日期范围_米易.Show_ME(str开始日期, str结束日期) = False Then Exit Sub
    '本地提取帐户支付总额
    gstrSQL = "Select SUM(发生费用金额) 发生费用金额,SUM(首先自付金额) 首先自付金额 ,SUM(个人帐户支付) 个人帐户支付" & _
        " From 保险结算记录 " & _
        " Where 性质=1 " & _
        " And 结算时间 Between [1] And [2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "统计门诊结算", CDate(str开始日期), CDate(str结束日期))
    If Not rsAccount.EOF Then
        cur发生费用金额 = Nvl(rsAccount!发生费用金额, 0)
        cur首先自付金额 = Nvl(rsAccount!首先自付金额, 0)
        cur个人帐户支付 = Nvl(rsAccount!个人帐户支付, 0)
    End If
    
    If Not 启动 Then Exit Sub
    '获取指定日期范围内的开始、结束就诊编号
    If Not FUNC_就诊编号(str开始日期, str结束日期, str开始就诊编号, str结束就诊编号, _
    str开始结算编号, str结束结算编号, False) Then Exit Sub
    '获取医保中心返回的统筹支付与统筹自付总额
    gstrPara_米易 = GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & _
        GetParaCode(就诊编号_起始, str开始就诊编号) & GetParaCode(就诊编号_截止, str结束就诊编号) '& _
        GetParaCode(结算编号_起始, str开始结算编号) & GetParaCode(结算编号_截止, str结束结算编号)
    Call 调用接口_米易("checkexpense")
    Call 医保终止_米易
    
    If Not (Format(cur发生费用金额, "#####0.00;-#####0.00;0;") = Format(gComInfo_米易.核对_医疗费总额, "#####0.00;-#####0.00;0;") _
    And Format(cur首先自付金额, "#####0.00;-#####0.00;0;") = Format(gComInfo_米易.核对_挂钩自费总额, "#####0.00;-#####0.00;0;") _
    And Format(cur个人帐户支付, "#####0.00;-#####0.00;0;") = Format(gComInfo_米易.核对_帐户支付总额, "#####0.00;-#####0.00;0;")) Then
        MsgBox "（本地）医疗费总额：" & cur发生费用金额 & String(4, " ") & "（医保）医疗费总额：" & gComInfo_米易.核对_医疗费总额 & vbCrLf & _
               "（本地）首先自付总额：" & cur首先自付金额 & String(4, " ") & "（医保）首先自付总额：" & gComInfo_米易.核对_挂钩自费总额 & vbCrLf & _
               "（本地）个人帐户支付总额：" & cur个人帐户支付 & String(4, " ") & "（医保）个人帐户支付总额：" & gComInfo_米易.核对_帐户支付总额
    Else
        MsgBox "数据正确无误，核对成功！", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub 核对住院结算_米易()
    Dim cur统筹支付总额 As Currency, cur统筹自付总额 As Currency
    Dim str开始日期 As String, str结束日期 As String
    Dim str开始就诊编号 As String, str结束就诊编号 As String
    Dim str开始结算编号 As String, str结束结算编号 As String
    Dim rsAccount As New ADODB.Recordset
    
    On Error GoTo errHand
    '获取查询日期
    If frm日期范围_米易.Show_ME(str开始日期, str结束日期) = False Then Exit Sub
    '本地提取帐户支付总额
    gstrSQL = "Select SUM(统筹报销金额) 统筹支付金额,SUM(进入统筹金额-统筹报销金额) 统筹自付金额 " & _
        " From 保险结算记录 " & _
        " Where 性质=2 " & _
        " And 结算时间 Between [1] And [2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "统计统筹支付总额", CDate(str开始日期), CDate(str结束日期))
    If Not rsAccount.EOF Then
        cur统筹支付总额 = Nvl(rsAccount!统筹支付金额, 0)
        cur统筹自付总额 = Nvl(rsAccount!统筹自付金额, 0)
    End If
    '获取指定日期范围内的开始、结束就诊编号
    If Not FUNC_就诊编号(str开始日期, str结束日期, str开始就诊编号, str结束就诊编号, _
    str开始结算编号, str结束结算编号, True) Then Exit Sub
    
    If Not 启动 Then Exit Sub
    '获取医保中心返回的统筹支付与统筹自付总额
    gstrPara_米易 = GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & _
        GetParaCode(就诊编号_起始, str开始就诊编号) & GetParaCode(就诊编号_截止, str结束就诊编号) & _
        GetParaCode(结算编号_起始, str开始结算编号) & GetParaCode(结算编号_截止, str结束结算编号)
    Call 调用接口_米易("checkexpensereckoninginfo")
    Call 医保终止_米易
    
    If Not (Format(cur统筹支付总额, "#####0.00;-#####0.00;0;") = Format(gComInfo_米易.核对_统筹支付金额, "#####0.00;-#####0.00;0;") _
    And Format(cur统筹自付总额, "#####0.00;-#####0.00;0;") = Format(gComInfo_米易.核对_统筹自付金额, "#####0.00;-#####0.00;0;")) Then
        MsgBox "（本地）统筹支付总额：" & cur统筹支付总额 & String(4, " ") & "（医保）统筹支付总额：" & gComInfo_米易.核对_统筹支付金额 & vbCrLf & _
               "（本地）统筹自付总额：" & cur统筹自付总额 & String(4, " ") & "（医保）统筹自付总额：" & gComInfo_米易.核对_统筹自付金额
    Else
        MsgBox "数据正确无误，核对成功！", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function FUNC_就诊编号(ByVal str开始日期 As String, ByVal str结束日期 As String, _
    str开始就诊编号 As String, str结束就诊编号 As String, _
    str开始结算编号 As String, str结束结算编号 As String, ByVal bln结算 As Boolean) As Boolean
    Dim rs顺序号 As New ADODB.Recordset
    '顺序号保存就诊编号，备注中保存当次的结算编号
    gstrSQL = "Select Min(A.支付顺序号) 开始就诊编号,Max(A.支付顺序号) 结束就诊编号 "
    If bln结算 Then gstrSQL = gstrSQL & ",Min(A.备注) 开始结算编号,Max(A.备注) 结束结算编号"
    gstrSQL = gstrSQL & _
             " From 保险结算记录 A" & _
             " Where A.险类=[1] And 结算时间 between [2] And [3]"
    Set rs顺序号 = zlDatabase.OpenSQLRecord(gstrSQL, "获取就诊编号", type_米易, CDate(str开始日期), CDate(str结束日期))
    
    If IsNull(rs顺序号!开始就诊编号) Then Exit Function
    
    str开始就诊编号 = rs顺序号!开始就诊编号
    str结束就诊编号 = rs顺序号!结束就诊编号
    If bln结算 Then
        str开始结算编号 = rs顺序号!开始结算编号
        str结束结算编号 = rs顺序号!结束结算编号
    End If
    FUNC_就诊编号 = True
End Function

Public Sub 补充上传门诊明细()
    On Error GoTo errHand
    Dim bln特种病 As Boolean, int上传 As Integer, blnError As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If Not 启动 Then Exit Sub
    If Not 调用接口_米易("getsysdate") Then Exit Sub
    
    gstrSQL = "Select A.ID,A.病人ID,A.NO,A.序号,A.记录性质,A.记录状态,A.登记时间,A.开单人 as 医生," & _
            "   A.数次*A.付数 as 数量,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额," & _
            "   A.收费类别,B.编码 as 项目编码,B.名称 as 项目名称,D.项目编码 医保编码,Nvl(H.类别,0) 特病种," & _
            "   C.名称 开单部门,E.名称 受单部门,A.摘要 记账流水号,F.支付顺序号 就诊编号,F.备注 结算编号,G.医保号 个人编号" & _
            " From (Select * From 门诊费用记录 Where 记录性质=1 And Nvl(是否上传,0)=0 And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0) A,收费细目 B,部门表 C,保险支付项目 D,部门表 E,保险结算记录 F,保险帐户 G,(Select * From 保险病种 Where 险类=[1]) H " & _
            " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID(+) And A.执行部门ID=E.ID(+) And A.收费细目ID=D.收费细目ID And D.险类=[1]" & _
            " And A.结帐ID=F.记录ID And F.性质=1 And A.病人ID=G.病人ID And G.险类=D.险类 And G.病种ID=H.ID(+)" & _
            " Order by A.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取本次结帐费用明细", type_米易)
    With rsTemp
        Do While Not .EOF
            '支付类别与医保病种是否是特种病有关
            bln特种病 = (!特病种 <> 0)
            gComInfo_米易.支付类别 = 201 'IIf(bln特种病, "0205", "0201")
            gstrPara_米易 = GetParaCode(个人编号, !个人编号) & GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & _
                GetParaCode(就诊编号, !就诊编号) & GetParaCode(记账流水号, !记账流水号) & _
                GetParaCode(结算编号, !结算编号) & GetParaCode(支付类别, gComInfo_米易.支付类别) & _
                GetParaCode(医保编码, !医保编码) & GetParaCode(项目编码, !项目编码) & GetParaCode(项目名称, !项目名称) & _
                GetParaCode(数量, !数量) & GetParaCode(费用总额, !结帐金额) & GetParaCode(开单科室, Nvl(!开单部门, "")) & _
                GetParaCode(开单医生, Nvl(!医生, "")) & GetParaCode(受单科室, Nvl(!受单部门, "")) & GetParaCode(受单医生, "") & _
                GetParaCode(经办时间, gComInfo_米易.系统时间)
            
            If 调用接口_米易("recipeinfotran") Then
                int上传 = 1
            Else
                int上传 = 0
                blnError = True
            End If
            
            '为病人费用记录打上标记，以便随时上传
            'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
            gstrSQL = "ZL_病人费用记录_更新医保(" & rsTemp("ID") & ",NULL,NULL,NULL,NULL," & int上传 & ",'" & !记账流水号 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
            .MoveNext
        Loop
    End With

    Call 医保终止_米易
    If blnError Then
        MsgBox "部分费用明细未正确上传，请到保险帐户管理中重新上传！", vbInformation, gstrSysName
    Else
        MsgBox "明细上传成功！", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Function 处方登记_米易(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '先写入单据头，再写入单据体
    '记录状态（1-新增;否则为删除），费用处单据只能整张单据删除后，再产生新单据
    On Error GoTo errHand
    处方登记_米易 = False

    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        gstrSQL = " Select A.病人ID,A.NO,A.序号,A.记录性质,A.记录状态,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 登记时间," & _
                  " A.开单人 医生,B.名称 开单部门,A.收费细目ID,C.项目编码 医保项目编码 ,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传" & _
                  " From 住院费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=[1]) C " & _
                  " Where A.记录性质=[2] And A.记录状态=[3] And A.NO=[4]" & _
                  " And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0" & _
                  " Order by A.病人ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "处方登记", type_米易, lng记录性质, lng记录状态, str单据号)
        If .RecordCount = 0 Then
            MsgBox "未找到处方记录，向医保服务器传输数据失败！[处方登记]", vbInformation, gstrSysName
            Exit Function
        End If
    End With

    If Not 启动 Then Exit Function
    If Not 上传处方_米易(rsTemp) Then Exit Function
    Call 医保终止_米易

    处方登记_米易 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function 上传处方_米易(ByVal rsExse As ADODB.Recordset) As Boolean
    Dim lng病人ID As Long
    Dim curTotal As Currency
    Dim blnUpload As Boolean, blnInsure As Boolean
    Dim rsTemp As New ADODB.Recordset, rsInsure As New ADODB.Recordset
    
    If Not 调用接口_米易("getsysdate") Then Exit Function
    
    gComInfo_米易.结算编号 = 0               '上传费用明细时，结算编号必需要置为零
    gComInfo_米易.支付类别 = "0301"
    gComInfo_米易.记账流水号 = GetSequence(gint记帐流水号)
    
    With rsExse
        Do While Not .EOF
            If IsNull(!医保项目编码) Then
                MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
                Exit Function
            End If
            .MoveNext
        Loop
        .MoveFirst
        
        '上传费用明细
        Do While Not .EOF
            If lng病人ID <> !病人ID Then
                '检查本次是否以医保身份入院
                gstrSQL = "Select Count(*) Records From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.病人ID=[1] And A.主页ID=B.住院次数 And A.险类=[2]"
                Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否医保病人", CLng(!病人ID), type_米易)
                blnInsure = (rsInsure!Records = 1)
                If blnInsure Then
                    blnInsure = 获取病人相关信息(!病人ID)
                    If blnInsure Then lng病人ID = !病人ID
                End If
            End If
            
            If blnInsure Then
                blnUpload = False
                If Not IsNull(!是否上传) Then
                    blnUpload = (!是否上传 = 0)
                End If
                
                If blnUpload Then
                    
                    '取该收费细目的编码与名称
                    gstrSQL = "Select 编码 项目编码 ,名称 项目名称 From 收费细目 Where ID = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该收费细目的编码与名称", CLng(!收费细目ID))
                    
                    gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & _
                        GetParaCode(就诊编号, gComInfo_米易.就诊编号) & GetParaCode(记账流水号, gComInfo_米易.记账流水号 & .AbsolutePosition) & _
                        GetParaCode(结算编号, gComInfo_米易.结算编号) & GetParaCode(支付类别, gComInfo_米易.支付类别) & _
                        GetParaCode(医保编码, !医保项目编码) & GetParaCode(项目编码, rsTemp!项目编码) & GetParaCode(项目名称, rsTemp!项目名称) & _
                        GetParaCode(数量, Abs(!数量)) & GetParaCode(费用总额, !金额) & GetParaCode(开单科室, Nvl(!开单部门, "")) & _
                        GetParaCode(开单医生, Nvl(!医生, "")) & GetParaCode(受单科室, "") & GetParaCode(受单医生, "") & _
                        GetParaCode(经办时间, Format(!登记时间, "yyyy-MM-dd HH:mm:ss"))
                    If 调用接口_米易("recipeinfotran") Then
                        '仅打上传标志，因为明细必须正确上传后，才能保证结算的正确性
                        'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                        gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
                    Else
                        Exit Function
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    
    上传处方_米易 = True
End Function


