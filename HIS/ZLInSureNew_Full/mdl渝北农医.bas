Attribute VB_Name = "mdl渝北农医"
Option Explicit
Public gstrBusiness_渝北农医 As String
Public gstrInput_渝北农医 As String
Public gstrOutput_渝北农医 As String

Private Const mstrAmountFormat As String = "#0.0000;-#0.0000;0;"
Private Const mstrPriceFormat As String = "#0.0000;-#0.0000;0;"
Private Const mstrDateFormat As String = "yyyy-MM-dd HH:mm:ss"
Private Const gstrSplit_渝北农医 As String = "|"
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Public Enum Business_渝北农医
    读取个人信息_渝北农医 = 300
    就诊登记_渝北农医 = 101
    就诊信息修改_渝北农医 = 102
    就诊登记取消_渝北农医 = 103
    处方明细上传_渝北农医 = 104
    处方明细作废_渝北农医 = 105
    预结算_渝北农医 = 106
    正式结算_渝北农医 = 107
    结算作废_渝北农医 = 108
End Enum

Private Type ComInfo_渝北农医
    医院编码 As String
    医院名称 As String
    业务类型 As String
    医疗证号 As String
    个人编号 As String
    就诊流水号 As String
    结算流水号 As String
    疾病编码 As String                      '保存身份验证后返回的疾病编码
    并发症 As String
    总费用 As Currency                      'HIS
    总费用_中心 As Currency                 '中心的费用总额
    门诊结算入参 As String
End Type
Public gComInfo_渝北农医 As ComInfo_渝北农医

Private gobjYB As Object   '定义存放引用对象的变量。
Private mblnInit As Boolean
Private strFields As String, strValues As String
Private mrsOutExse As New ADODB.Recordset

Public Function 身份标识_渝北农医(Optional bytType As Byte, Optional lng病人ID As Long) As String
    Dim StrInput As String
    Dim strIdentify As String
    Dim strRegistCode As String             '挂号单号
    Dim strRegisterOffice As String         '就诊科室
    Dim strRegisterDoctor As String         '医生
    Dim rsTemp As New ADODB.Recordset
    Dim strDate As String
    On Error GoTo errHand
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    strIdentify = frmIdentify渝北农医.GetPatient(bytType, lng病人ID)
    If strIdentify = "" Then Exit Function
    If Not (bytType = 1 Or bytType = 0 Or bytType = 3) Then Exit Function
    
    '进行门诊登记
    If bytType = 0 Then
        '入参：合作医疗号码│合作医疗病人在医院就诊的挂号号码│就诊的医疗类别│医院就诊的科室│就诊的医生│" & _
        医院的诊断│医院就诊登记的日期│并发症│就诊机构的机构编码│就诊机构的机构名称│经办单位│经办人
        '取当天挂号的科室与医生
        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        gstrSQL = " Select B.名称 AS 挂号科室,执行人 AS 医生 " & _
                  " From 门诊费用记录 A,部门表 B " & _
                  " Where A.记录性质=4 And 记录状态=1 And 病人ID=[1]" & _
                  " And A.执行部门ID=B.ID And 登记时间 Between [2] And [3] And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当天挂号的科室与医生", lng病人ID, CDate(strDate), CDate(strDate & " 23:59:59"))
        If rsTemp.RecordCount = 0 Then
            MsgBox "今天没有有效的挂号记录,无法进行门诊就诊登记！", vbInformation, gstrSysName
            Exit Function
        End If
        strRegisterOffice = Nvl(rsTemp!挂号科室)
        strRegisterDoctor = Nvl(rsTemp!医生)
        
        '获取挂号单号，十位，唯一标识
        strRegistCode = Right(CStr(zlDatabase.GetNextID("部门表")), 10)
        StrInput = gComInfo_渝北农医.医疗证号 & gstrSplit_渝北农医 & strRegistCode & gstrSplit_渝北农医 & _
            gComInfo_渝北农医.业务类型 & gstrSplit_渝北农医 & strRegisterOffice & gstrSplit_渝北农医 & _
            strRegisterDoctor & gstrSplit_渝北农医 & gComInfo_渝北农医.疾病编码 & gstrSplit_渝北农医 & _
            Format(zlDatabase.Currentdate(), mstrDateFormat) & gstrSplit_渝北农医 & gComInfo_渝北农医.并发症 & gstrSplit_渝北农医 & _
            gComInfo_渝北农医.医院编码 & gstrSplit_渝北农医 & gComInfo_渝北农医.医院名称 & gstrSplit_渝北农医 & _
            gComInfo_渝北农医.医院编码 & gstrSplit_渝北农医 & UserInfo.姓名
        Call 调用接口_准备_渝北农医(就诊登记_渝北农医, StrInput)
        If Not 调用接口_渝北农医() Then Exit Function
        
        '更新流水号
        gComInfo_渝北农医.就诊流水号 = Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(1)
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_渝北农医 & ",'顺序号','''" & gComInfo_渝北农医.就诊流水号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊流水号")
        
        '初始化记录集
        strFields = "就诊流水号," & adLongVarChar & ",100" & gstrSplit_渝北农医 & "明细流水号," & adLongVarChar & ",20" & gstrSplit_渝北农医 & _
            "开方日期," & adLongVarChar & ",20" & gstrSplit_渝北农医 & "医保编码," & adLongVarChar & ",50" & gstrSplit_渝北农医 & _
            "项目名称," & adLongVarChar & ",100" & gstrSplit_渝北农医 & "规格," & adLongVarChar & ",100" & gstrSplit_渝北农医 & _
            "剂型," & adLongVarChar & ",100" & gstrSplit_渝北农医 & "单价," & adLongVarChar & ",20" & gstrSplit_渝北农医 & _
            "数量," & adLongVarChar & ",20" & gstrSplit_渝北农医 & "上传流水号," & adLongVarChar & ",20"
        Call Record_Init(mrsOutExse, strFields)
    End If
    
    '更新保险帐户相关信息（统筹区号、业务类型）
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_渝北农医 & ",'业务类型','''" & gComInfo_渝北农医.业务类型 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存业务类型")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_渝北农医 & ",'并发症','''" & gComInfo_渝北农医.并发症 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存并发症")
    
    '返回病人信息串
    身份标识_渝北农医 = strIdentify
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub 取消就诊登记_渝北农医(Optional bytType As Byte, Optional lng病人ID As Long)
    '取消本次就诊登记，如果预结算时已上传处方明细，则先取消明细，再取消就诊登记
    If bytType <> 0 Then Exit Sub       '只可能是门诊或挂号
    On Error GoTo errHand
    
    '先作废上次上传的所有处方明细
    With mrsOutExse
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Call 调用接口_准备_渝北农医(处方明细作废_渝北农医, !上传流水号)
            Call 调用接口_渝北农医
            .MoveNext
        Loop
    End With
    
    '取消就诊登记
    Call 调用接口_准备_渝北农医(就诊登记取消_渝北农医, gComInfo_渝北农医.就诊流水号)
    Call 调用接口_渝北农医
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function 医保初始化_渝北农医(Optional ByVal blnTest As Boolean = False) As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    Dim strServer As String, strUser As String, strPass As String, strDatabase As String
    Dim rsTemp As New ADODB.Recordset
    Dim cnTest As New ADODB.Connection

    On Error Resume Next
    
    If mblnInit = False Then
        If Not blnTest Then '如果是测试，则说明是保险参数设置处调用
            '读出连接医保服务器的配置
            gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取保险参数", TYPE_渝北农医)
            
            Do Until rsTemp.EOF
                Select Case rsTemp("参数名")
                    Case "医保用户名"
                        strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Case "医保服务器"
                        strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Case "医保用户密码"
                        strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Case "医保实例名"
                        strDatabase = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                End Select
                rsTemp.MoveNext
            Loop
            
            If OpenSQLServer(cnTest, strServer, strUser, strPass, strDatabase) = False Then
                MsgBox "无法连接到前置机，请检查保险参数是否设置正确！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        Set gobjYB = CreateObject("HisSel.Handld")
        '检查连接是否建立
        If gobjYB Is Nothing Then
            MsgBox "医保初始化失败！", vbInformation, gstrSysName
            '调试重庆医保银海版 204-04-07
            Exit Function
        End If
        '取医院编码
        gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_渝北农医)
        gComInfo_渝北农医.医院编码 = Nvl(rsTemp!医院编码)
        '取医院名称
        gstrSQL = "Select JGMC 医院名称 From JGDJ Where JGBM='" & gComInfo_渝北农医.医院编码 & "'"
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open gstrSQL, cnTest
        gComInfo_渝北农医.医院名称 = Nvl(rsTemp!医院名称)
        
        cnTest.Close
        If Not blnTest Then mblnInit = True
    End If
    
    医保初始化_渝北农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保设置_渝北农医() As Boolean
    医保设置_渝北农医 = frmSet渝北农医.参数设置
End Function

Public Function 医保终止_渝北农医() As Boolean
    On Error Resume Next
    
    Set gobjYB = Nothing
    mblnInit = False
    医保终止_渝北农医 = True
End Function

Public Sub 调用接口_准备_渝北农医(ByVal strBusiness As String, Optional ByVal StrInput As String = "")
    gstrBusiness_渝北农医 = strBusiness
    gstrInput_渝北农医 = StrInput
End Sub

Public Function 调用接口_渝北农医() As Boolean
    Dim arrOutput
    Dim lngResult As Long
    On Error GoTo errHand
    
    Call gobjYB.Business(gstrBusiness_渝北农医, gstrInput_渝北农医, gstrOutput_渝北农医)
    Call WriteInfo(String(20, "-"))
    Call WriteInfo("交易号：" & gstrBusiness_渝北农医)
    Call WriteInfo("入参：" & gstrInput_渝北农医)
    Call WriteInfo("出参：" & gstrOutput_渝北农医)
    
    arrOutput = Split(gstrOutput_渝北农医, gstrSplit_渝北农医)
    lngResult = Val(arrOutput(0))
    If lngResult < 0 Then               '错误信息
        MsgBox "交易类型[" & gstrBusiness_渝北农医 & "]错误代码[" & lngResult & "]" & arrOutput(1), vbInformation, gstrSysName
        Exit Function
    End If
    
    调用接口_渝北农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_渝北农医(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim StrInput As String
    Dim lng病人ID As Long
    Dim str开方日期 As String, str处方号 As String, str医保编码 As String, str项目名称 As String, str规格 As String, str剂型 As String
    Dim dbl帐户支付 As Double, dbl现金 As Double, dbl优惠金额 As Double
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    On Error GoTo errHand
    
    lng病人ID = rs明细!病人ID
    str开方日期 = Format(zlDatabase.Currentdate, mstrDateFormat)
    '先作废上次上传的所有处方明细
    With mrsOutExse
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            StrInput = !上传流水号
            Call 调用接口_准备_渝北农医(处方明细作废_渝北农医, StrInput)
            If Not 调用接口_渝北农医() Then Exit Function
            .MoveNext
        Loop
    End With
    
    '获取该病人的就诊时间
    gstrSQL = "Select to_char(就诊时间,'yyyy-MM-dd hh24:mi:ss') As 就诊时间 From 保险帐户" & _
        " Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取该病人的就诊时间", TYPE_渝北农医, lng病人ID)
    str处方号 = GetSequence(Format(rsTemp!就诊时间, "yyMMddHHmmss")) & Left(CStr(Rnd() * 100), 2)
    
    '初始化记录集
    strFields = "就诊流水号," & adLongVarChar & ",100" & gstrSplit_渝北农医 & "明细流水号," & adLongVarChar & ",20" & gstrSplit_渝北农医 & _
        "开方日期," & adLongVarChar & ",20" & gstrSplit_渝北农医 & "医保编码," & adLongVarChar & ",50" & gstrSplit_渝北农医 & _
        "项目名称," & adLongVarChar & ",100" & gstrSplit_渝北农医 & "规格," & adLongVarChar & ",100" & gstrSplit_渝北农医 & _
        "剂型," & adLongVarChar & ",100" & gstrSplit_渝北农医 & "单价," & adLongVarChar & ",20" & gstrSplit_渝北农医 & _
        "数量," & adLongVarChar & ",20" & gstrSplit_渝北农医 & "上传流水号," & adLongVarChar & ",20"
    Call Record_Init(mrsOutExse, strFields)
        
    '得到本次结算的总费用
    strFields = "就诊流水号" & gstrSplit_渝北农医 & "明细流水号" & gstrSplit_渝北农医 & _
            "开方日期" & gstrSplit_渝北农医 & "医保编码" & gstrSplit_渝北农医 & _
            "项目名称" & gstrSplit_渝北农医 & "规格" & gstrSplit_渝北农医 & _
            "剂型" & gstrSplit_渝北农医 & "单价" & gstrSplit_渝北农医 & "数量" & gstrSplit_渝北农医 & "上传流水号"
    With rs明细
        If .RecordCount > 99 Then
            MsgBox "门诊处方明细不能超过99条记录！", vbInformation, gstrSysName
            Exit Function
        End If
        '求费用总额
        gComInfo_渝北农医.总费用 = 0
        Do While Not .EOF
            gComInfo_渝北农医.总费用 = gComInfo_渝北农医.总费用 + !实收金额
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        Do While Not .EOF
            '提取收费细目的相关信息
            gstrSQL = " Select A.类别 AS 收费类别,A.名称,A.规格,B.项目编码 From 收费细目 A,保险支付项目 B" & _
                      " Where A.ID=B.收费细目ID(+) And B.险类(+)=[1] And A.ID=[2]"
            Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目信息", TYPE_渝北农医, CLng(!收费细目ID))
            str项目名称 = Nvl(rsItem!名称)
            str医保编码 = Nvl(rsItem!项目编码)
            str规格 = Nvl(rsItem!规格)
            If InStr(1, str规格, "|") <> 0 Then str规格 = Mid(str规格, 1, InStr(1, str规格, "|") - 1)
            
            '如果是药品，取剂型
            str剂型 = ""
            If InStr(1, "5,6,7", rsItem!收费类别) <> 0 Then
                gstrSQL = "SELECT 名称 FROM 药品剂型 WHERE 编码=(SELECT 剂型 FROM 药品信息 WHERE 药名ID=(SELECT 药名ID FROM 药品目录 WHERE 药品ID=[1]))"
                Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "提取药品的剂型", CLng(!收费细目ID))
                str剂型 = Nvl(rsItem!名称)
            End If
            
            '处方明细上传入参：合作医疗病人的就诊流水号|医院开出的处方的编码|医院开处方的日期|药品或诊疗项目合作医疗对照的编码| & _
            药品或者诊疗项目的名称|药品的规格|药品的剂型|药品或者诊疗项目的单价|药品数量或者诊疗次数
            StrInput = gComInfo_渝北农医.就诊流水号 & gstrSplit_渝北农医 & str处方号 & String(2 - Len(CStr(.AbsolutePosition)), "0") & .AbsolutePosition & gstrSplit_渝北农医 & _
                str开方日期 & gstrSplit_渝北农医 & str医保编码 & gstrSplit_渝北农医 & Left(str项目名称, 15) & gstrSplit_渝北农医 & _
                Left(str规格, 10) & gstrSplit_渝北农医 & Left(str剂型, 10) & gstrSplit_渝北农医 & Format(!单价, mstrPriceFormat) & gstrSplit_渝北农医 & Format(!数量, mstrAmountFormat)
            
            Call 调用接口_准备_渝北农医(处方明细上传_渝北农医, StrInput)
            If Not 调用接口_渝北农医() Then Exit Function
            
            '将已上传的处方明细写入记录集
            strValues = gComInfo_渝北农医.就诊流水号 & gstrSplit_渝北农医 & str处方号 & String(2 - Len(CStr(.AbsolutePosition)), "0") & .AbsolutePosition & gstrSplit_渝北农医 & _
                str开方日期 & gstrSplit_渝北农医 & str医保编码 & gstrSplit_渝北农医 & str项目名称 & gstrSplit_渝北农医 & _
                str规格 & gstrSplit_渝北农医 & str剂型 & gstrSplit_渝北农医 & Format(!单价, mstrPriceFormat) & gstrSplit_渝北农医 & Format(!数量, mstrAmountFormat) & gstrSplit_渝北农医 & Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(1)
            Call Record_Add(mrsOutExse, strFields, strValues)
            
            .MoveNext
        Loop
    End With
    
    '预结算的入参：病人就诊登记的流水号|结算的类别|住院的床日|经办单位|经办人|经办日期
    StrInput = gComInfo_渝北农医.就诊流水号 & gstrSplit_渝北农医 & "01" & gstrSplit_渝北农医 & "0" & gstrSplit_渝北农医 & _
        gComInfo_渝北农医.医院编码 & gstrSplit_渝北农医 & UserInfo.姓名 & gstrSplit_渝北农医 & str开方日期
    gComInfo_渝北农医.门诊结算入参 = StrInput
    Call 调用接口_准备_渝北农医(预结算_渝北农医, StrInput)
    If Not 调用接口_渝北农医() Then Exit Function
    
    '出参：执行代码│结算流水号│这次结算医院总的金额│经过医院下浮后的总金额│合作医疗办公室承认可以参加报销的金额│实际报销的金额│病人自负的金额
    dbl优惠金额 = Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(2)) - Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(3))
    dbl帐户支付 = Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(5))
    dbl现金 = Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(6))
    
    str结算方式 = "家庭帐户;" & dbl帐户支付 & ";0|优惠金额;" & dbl优惠金额 & ";0"
    门诊虚拟结算_渝北农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_渝北农医(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    Dim lng病人ID As Long
    Dim StrInput As String
    Dim str结算日期 As String, str结算流水号 As String, str就诊顺序号 As String
    Dim dbl进入统筹 As Double, dbl统筹报销 As Double, dbl现金 As Double, dbl优惠金额 As Double
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '预结算的入参：病人就诊登记的流水号|结算的类别|住院的床日|经办单位|经办人|经办日期,结算日期
    str结算日期 = Format(zlDatabase.Currentdate, mstrDateFormat)
    StrInput = gComInfo_渝北农医.门诊结算入参 & gstrSplit_渝北农医 & str结算日期
    Call 调用接口_准备_渝北农医(正式结算_渝北农医, StrInput)
    If Not 调用接口_渝北农医() Then Exit Function
    
    '出参：执行代码│结算流水号│这次结算医院总的金额│经过医院下浮后的总金额│合作医疗办公室承认可以参加报销的金额│实际报销的金额│病人自负的金额
    str结算流水号 = Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(1)
    dbl优惠金额 = Format(Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(2)) - Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(3)), "#0.00")
    dbl进入统筹 = Format(Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(4)), "#0.00")
    dbl统筹报销 = Format(Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(5)), "#0.00")   '统筹报销就是家庭帐户允许支付额
    dbl现金 = Format(Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(6)), "#0.00")
    
    '取病人ID
    gstrSQL = "Select 病人ID From 门诊费用记录 Where 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该病人的ID", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    
    '取就诊顺序号
    gstrSQL = "Select 顺序号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取就诊顺序号", TYPE_渝北农医, lng病人ID)
    str就诊顺序号 = Nvl(rsTemp!顺序号)
    
    '保存本次结算情况
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_渝北农医 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_渝北农医.总费用 & "," & dbl现金 & "," & dbl优惠金额 & "," & dbl进入统筹 & "," & dbl统筹报销 & ",0,0," & _
        dbl统筹报销 & ",'" & str就诊顺序号 & "|" & str结算流水号 & "',null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊收费数据")
    
    门诊结算_渝北农医 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_渝北农医(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    Dim lng冲销ID As Long
    Dim str结算流水号 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    '取冲销记录的结帐ID，单据号
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!结帐ID
    
    '取结算流水号
    gstrSQL = "Select * From 保险结算记录 Where 性质=1 And 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取结算流水号", lng结帐ID)
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "没有找到原始结算记录，无法进行门诊结算冲销！"
        Exit Function
    End If
    gComInfo_渝北农医.就诊流水号 = Split(rsTemp!支付顺序号, gstrSplit_渝北农医)(0)
    str结算流水号 = Split(rsTemp!支付顺序号, gstrSplit_渝北农医)(1)
    
    '调用结算冲销
    Call 调用接口_准备_渝北农医(结算作废_渝北农医, str结算流水号)
    If Not 调用接口_渝北农医() Then Exit Function
    
    '取消就诊登记
    Call 调用接口_准备_渝北农医(就诊登记取消_渝北农医, gComInfo_渝北农医.就诊流水号)
    Call 调用接口_渝北农医
    
    '保存本次结算情况
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_渝北农医 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & "," & -1 * Nvl(rsTemp!首先自付金额, 0) & "," & -1 * Nvl(rsTemp!进入统筹金额, 0) & "," & -1 * Nvl(rsTemp!统筹报销金额, 0) & ",0,0," & _
        -1 * Nvl(rsTemp!个人帐户支付, 0) & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "门诊结算冲销")
    
    门诊结算冲销_渝北农医 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 入院登记_渝北农医(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    Dim StrInput As String
    Dim strRegistCode As String             '挂号单号
    Dim strInHospitalDate As String         '入院日期
    Dim strRegisterOffice As String         '就诊科室
    Dim strRegisterDoctor As String         '医生
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '取科室与医生
    gstrSQL = " Select A.入院日期,B.名称 科室,A.住院医师 医生 From 病案主页 A,部门表 B " & _
              " Where A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取科室与医生", lng病人ID, lng主页ID)
    strInHospitalDate = Format(rsTemp!入院日期, mstrDateFormat)
    strRegisterDoctor = Nvl(rsTemp!医生)
    strRegisterOffice = Nvl(rsTemp!科室)
    
    '入参：合作医疗号码│合作医疗病人在医院就诊的挂号号码│就诊的医疗类别│医院就诊的科室│就诊的医生│" & _
    医院的诊断│医院就诊登记的日期│并发症│就诊机构的机构编码│就诊机构的机构名称│经办单位│经办人
    '获取挂号单号，十位，唯一标识
    strRegistCode = Right(CStr(zlDatabase.GetNextID("部门表")), 10)
    StrInput = gComInfo_渝北农医.医疗证号 & gstrSplit_渝北农医 & strRegistCode & gstrSplit_渝北农医 & _
        gComInfo_渝北农医.业务类型 & gstrSplit_渝北农医 & strRegisterOffice & gstrSplit_渝北农医 & _
        strRegisterDoctor & gstrSplit_渝北农医 & gComInfo_渝北农医.疾病编码 & gstrSplit_渝北农医 & _
        strInHospitalDate & gstrSplit_渝北农医 & gComInfo_渝北农医.并发症 & gstrSplit_渝北农医 & _
        gComInfo_渝北农医.医院编码 & gstrSplit_渝北农医 & gComInfo_渝北农医.医院名称 & gstrSplit_渝北农医 & _
        gComInfo_渝北农医.医院编码 & gstrSplit_渝北农医 & UserInfo.姓名
    Call 调用接口_准备_渝北农医(就诊登记_渝北农医, StrInput)
    If Not 调用接口_渝北农医() Then Exit Function
    
    '更新流水号
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_渝北农医 & ",'顺序号','''" & Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(1) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊流水号")
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_渝北农医 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    入院登记_渝北农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_渝北农医(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If 存在未结费用(lng病人ID, lng主页ID) Then
        MsgBox "该医保病人存在未结费用，不允许办理撤销入院登记！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '获取原就诊流水号
    gstrSQL = "Select 顺序号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取就诊流水号", TYPE_渝北农医, lng病人ID)
    gComInfo_渝北农医.就诊流水号 = rsTemp!顺序号
    
    '调用就诊登记作废接口
    Call 调用接口_准备_渝北农医(就诊登记取消_渝北农医, gComInfo_渝北农医.就诊流水号)
    If Not 调用接口_渝北农医 Then Exit Function
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_渝北农医 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    
    入院登记撤销_渝北农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_渝北农医(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHand
    
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_渝北农医 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    
    出院登记_渝北农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_渝北农医(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHand
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_渝北农医 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销出院登记")
    出院登记撤销_渝北农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 个人余额_渝北农医(strSelfNo As String) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '功能: 提取参保病人个人帐户余额
    '参数: strSelfNO-病人个人编号
    '返回: 返回个人帐户余额的金额
    '如果是门诊，返回家庭帐户余额；住院返回个人帐户余额
    gstrSQL = "Select Nvl(帐户余额,0) AS 个人帐户,Nvl(家庭帐户余额,0) AS 家庭帐户,病人ID From 保险帐户 Where 医保号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人ID", strSelfNo)
    个人余额_渝北农医 = rsTemp!个人帐户 + rsTemp!家庭帐户
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 处方上传_渝北农医(ByVal int性质 As Integer, ByVal int状态 As Integer, ByVal strNO As String) As Boolean
    Dim lng费用ID As Long, lng病人ID As Long
    Dim StrInput As String
    Dim blnInsure As Boolean
    Dim str就诊流水号 As String, str处方号 As String
    Dim str项目名称 As String, str医保编码 As String, str规格 As String, str剂型 As String
    Dim rsDetail As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    On Error GoTo errHand
    '上传处方明细
    '打开本次待上传的处方明细
    gstrSQL = " Select A.ID,A.记录性质,A.记录状态,A.NO,A.序号,A.收费类别,A.病人ID,A.收费细目ID,A.登记时间,Nvl(A.付数,1)*数次 AS 数量,A.实收金额/(Nvl(A.付数,1)*A.数次) AS 价格" & _
              " From 住院费用记录 A,保险帐户 B" & _
              " Where A.记录性质=[1] ANd A.记录状态=[2] And A.NO=[3] And Nvl(A.是否上传,0)=0" & _
              " And A.病人ID=B.病人ID And B.险类=[4]" & _
              " Order by 病人ID"
    Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次待上传的处方明细", int性质, int状态, strNO, TYPE_渝北农医)
    
    '先检查明细，不允许负数记帐（只针对医保病人的正常记帐的处方明细）
    With rsDetail
        lng病人ID = 0
        If int状态 = 1 Then
            Do While Not .EOF
                If lng病人ID <> !病人ID Then
                    lng病人ID = !病人ID
                    blnInsure = IsYBPatient(lng病人ID, str就诊流水号)
                End If
                If blnInsure Then
                    If !数量 < 0 Then
                        MsgBox "渝北农村合作医疗接口不支持为医保病人进行负数记帐，请直接冲销原始处方明细！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                .MoveNext
            Loop
        End If
        If .RecordCount <> 0 Then .MoveFirst
        
        '上传处方（上传几条算几条，没上传成功也允许保存单据）
        lng病人ID = 0
        Do While Not .EOF
            If lng病人ID <> !病人ID Then
                lng病人ID = !病人ID
                blnInsure = IsYBPatient(lng病人ID, str就诊流水号)
            End If
            
            If blnInsure Then
                '以费用ID后十位，做为本次处方明细流水号
                lng费用ID = !ID
                str处方号 = Right(CStr(lng费用ID), 10)
                
                '提取收费细目的相关信息
                gstrSQL = " Select A.类别 AS 收费类别,A.名称,A.规格,B.项目编码 From 收费细目 A,保险支付项目 B" & _
                          " Where A.ID=B.收费细目ID(+) And B.险类(+)=[1] And A.ID=[2]"
                Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目信息", TYPE_渝北农医, CLng(!收费细目ID))
                str项目名称 = Nvl(rsItem!名称)
                str医保编码 = Nvl(rsItem!项目编码)
                str规格 = Nvl(rsItem!规格)
                If InStr(1, str规格, "|") <> 0 Then str规格 = Mid(str规格, 1, InStr(1, str规格, "|") - 1)
                
                '如果是药品，取剂型
                str剂型 = ""
                If InStr(1, "5,6,7", rsItem!收费类别) <> 0 Then
                    gstrSQL = "SELECT 名称 FROM 药品剂型 WHERE 编码=(SELECT 剂型 FROM 药品信息 WHERE 药名ID=(SELECT 药名ID FROM 药品目录 WHERE 药品ID=[1]))"
                    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "提取药品的剂型", CLng(!收费细目ID))
                    str剂型 = Nvl(rsItem!名称)
                End If
                
                '处方明细上传入参：合作医疗病人的就诊流水号|医院开出的处方的编码|医院开处方的日期|药品或诊疗项目合作医疗对照的编码| & _
                药品或者诊疗项目的名称|药品的规格|药品的剂型|药品或者诊疗项目的单价|药品数量或者诊疗次数
                If int状态 <> 2 Then
                    StrInput = str就诊流水号 & gstrSplit_渝北农医 & str处方号 & gstrSplit_渝北农医 & _
                        Format(!登记时间, mstrDateFormat) & gstrSplit_渝北农医 & str医保编码 & gstrSplit_渝北农医 & Left(str项目名称, 15) & gstrSplit_渝北农医 & _
                        Left(str规格, 10) & gstrSplit_渝北农医 & Left(str剂型, 10) & gstrSplit_渝北农医 & Format(!价格, mstrPriceFormat) & gstrSplit_渝北农医 & Format(!数量, mstrAmountFormat)
                Else
                    '取原始费用记录ID
                    gstrSQL = "Select 摘要 From 住院费用记录 Where 记录性质=[1] And 记录状态=3 And NO=[2] And 序号=[3]"
                    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "取费用记录ID", CLng(!记录性质), CStr(!NO), CLng(!序号))
                    StrInput = Nvl(rsItem!摘要)
                    If Trim(StrInput) = "" Then
                        MsgBox "原始处方明细还未上传，销帐明细无法上传！", vbInformation, gstrSysName
                        处方上传_渝北农医 = True
                        Exit Function
                    End If
                End If
                
                Call 调用接口_准备_渝北农医(IIf(int状态 <> 2, 处方明细上传_渝北农医, 处方明细作废_渝北农医), StrInput)
                If Not 调用接口_渝北农医() Then
                    处方上传_渝北农医 = True
                    Exit Function
                End If
                
                '打上传标志
                If int状态 <> 2 Then
                    gstrSQL = "ZL_病人记帐记录_上传(" & lng费用ID & ",0,'" & Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(1) & "')"
                Else
                    gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
                End If
                Call zlDatabase.ExecuteProcedure(gstrSQL, "打上传标志")
            End If
            .MoveNext
        Loop
    End With
    
    处方上传_渝北农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_渝北农医(rsExse As Recordset, ByVal lng病人ID As Long) As String
    Dim intDO As Integer
    Dim lng费用ID As Long, lng主页ID As Long, lng住院天数 As Long
    Dim dbl帐户支付 As Double, dbl现金 As Double, dbl优惠金额 As Double
    Dim bln正常结算 As Boolean                  '正常结算(01)或转院结算(02)
    Dim StrInput As String
    Dim str处方号 As String, str就诊流水号 As String
    Dim str项目名称 As String, str医保编码 As String, str规格 As String, str剂型 As String
    Dim rsItem As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    On Error GoTo errHand
    
    '取出院方式
    gstrSQL = "Select 出院方式,住院天数,主页ID From 病案主页 Where (病人ID,主页ID) in (Select 病人ID,住院次数 From 病人信息 Where 病人ID=[1])"
    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "取出院方式", lng病人ID)
    lng主页ID = rsItem!主页ID
    lng住院天数 = Nvl(rsItem!住院天数, 0)
    bln正常结算 = IIf(rsItem!出院方式 = "转院", False, True)
    
    '获取病人的就诊流水号
    gstrSQL = "Select 顺序号 From 保险帐户 Where 险类=" & TYPE_渝北农医 & " And 病人ID=[1]"
    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人的就诊流水号", lng病人ID)
    str就诊流水号 = Nvl(rsItem!顺序号)
    
    '提取本次费用明细
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.收费类别,A.记录性质,A.记录状态,A.序号,A.收费细目ID,C.项目编码 AS 医保项目编码,B.编码,B.名称,A.实收金额 AS 金额" & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格,A.开单人 AS 医生,A.登记时间 " & _
              "  From 住院费用记录 A,收费细目 B,保险支付项目 C " & _
              "  where A.病人ID=[1] and A.主页ID=[2] and A.记帐费用=1 And A.操作员姓名 is not null AND A.实收金额 IS NOT NULL " & _
              "        and nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= [3]" & _
              "  Order by A.病人ID,A.发生时间"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次费用明细", lng病人ID, lng主页ID, TYPE_渝北农医)
    
    With rsExse
        '求费用总额
        gComInfo_渝北农医.总费用 = 0
        Do While Not .EOF
            gComInfo_渝北农医.总费用 = gComInfo_渝北农医.总费用 + !金额
            .MoveNext
        Loop
    End With
        
    With rs明细
        For intDO = 1 To 2
            .Filter = IIf(intDO = 1, "记录状态<>2", "记录状态=2")
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                '以费用ID后十位，做为本次处方明细流水号
                gstrSQL = "Select ID,实收金额 From 住院费用记录 Where 记录性质=[1] And 记录状态=[2] And NO=[3] And 序号=[]"
                Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "取费用记录ID", CLng(!记录性质), CLng(!记录状态), CStr(!NO), CLng(!序号))
                If Not IsNull(rsItem!实收金额) Then
                    lng费用ID = rsItem!ID
                    str处方号 = Right(rsItem!ID, 10)
                    
                    '提取收费细目的相关信息
                    gstrSQL = " Select A.类别 AS 收费类别,A.名称,A.规格,B.项目编码 From 收费细目 A,保险支付项目 B" & _
                              " Where A.ID=B.收费细目ID(+) And B.险类(+)=[1] And A.ID=[2]"
                    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目信息", TYPE_渝北农医, CLng(!收费细目ID))
                    str项目名称 = Nvl(rsItem!名称)
                    str医保编码 = Nvl(rsItem!项目编码)
                    str规格 = Nvl(rsItem!规格)
                    If InStr(1, str规格, "|") <> 0 Then str规格 = Mid(str规格, 1, InStr(1, str规格, "|") - 1)
                    
                    '如果是药品，取剂型
                    str剂型 = ""
                    If InStr(1, "5,6,7", rsItem!收费类别) <> 0 Then
                        gstrSQL = "SELECT 名称 FROM 药品剂型 WHERE 编码=(SELECT 剂型 FROM 药品信息 WHERE 药名ID=(SELECT 药名ID FROM 药品目录 WHERE 药品ID=[1]))"
                        Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "提取药品的剂型", CLng(!收费细目ID))
                        str剂型 = Nvl(rsItem!名称)
                    End If
                    
                    '处方明细上传入参：合作医疗病人的就诊流水号|医院开出的处方的编码|医院开处方的日期|药品或诊疗项目合作医疗对照的编码| & _
                    药品或者诊疗项目的名称|药品的规格|药品的剂型|药品或者诊疗项目的单价|药品数量或者诊疗次数
                    If intDO = 1 Then
                        StrInput = str就诊流水号 & gstrSplit_渝北农医 & str处方号 & gstrSplit_渝北农医 & _
                            Format(!登记时间, mstrDateFormat) & gstrSplit_渝北农医 & str医保编码 & gstrSplit_渝北农医 & Left(str项目名称, 15) & gstrSplit_渝北农医 & _
                            Left(str规格, 10) & gstrSplit_渝北农医 & Left(str剂型, 10) & gstrSplit_渝北农医 & Format(!价格, mstrPriceFormat) & gstrSplit_渝北农医 & Format(!数量, mstrAmountFormat)
                    Else
                        '取原始费用记录ID
                        gstrSQL = "Select 摘要 From 住院费用记录 Where 记录性质=[1] And 记录状态=3 And NO=[2] And 序号=[3]"
                        Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "取费用记录ID", CLng(!记录性质), CStr(!NO), CLng(!序号))
                        StrInput = Nvl(rsItem!摘要)
                    End If
                    
                    Call 调用接口_准备_渝北农医(IIf(intDO = 1, 处方明细上传_渝北农医, 处方明细作废_渝北农医), StrInput)
                    If Not 调用接口_渝北农医() Then Exit Function
                    
                    '打上传标志
                    If intDO = 1 Then
                        gstrSQL = "ZL_病人记帐记录_上传(" & lng费用ID & ",0,'" & Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(1) & "')"
                    Else
                        gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
                    End If
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "打上传标志")
                End If
                .MoveNext
            Loop
        Next
    End With
    
    '进行住院预结算，入参：病人就诊登记的流水号|结算类别|住院的床日|经办单位|经办人|经办日期
    StrInput = str就诊流水号 & gstrSplit_渝北农医 & IIf(bln正常结算, "01", "02") & gstrSplit_渝北农医 & _
        lng住院天数 & gstrSplit_渝北农医 & gComInfo_渝北农医.医院编码 & gstrSplit_渝北农医 & UserInfo.姓名 & gstrSplit_渝北农医 & Format(zlDatabase.Currentdate, mstrDateFormat)
    Call 调用接口_准备_渝北农医(预结算_渝北农医, StrInput)
    If Not 调用接口_渝北农医() Then Exit Function
    
    '出参：执行代码│结算流水号│这次结算医院总的金额│经过医院下浮后的总金额│合作医疗办公室承认可以参加报销的金额│实际报销的金额│病人自负的金额
    dbl优惠金额 = Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(2)) - Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(3))
    dbl帐户支付 = Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(5))
    dbl现金 = Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(6))
    
    '判断HIS总金额与医保返回的总金额是否一致
    If Format(gComInfo_渝北农医.总费用, "#####0.00") <> Format(Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(2)), "#####0.00") Then
        If MsgBox("医院总费用与合医办总费用不一致，是否继续？" & vbCrLf & _
            "医院总费用：" & Format(gComInfo_渝北农医.总费用, "#####0.00") & vbCrLf & _
            "中心总费用：" & Format(Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(2)), "#####0.00"), _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
        End If
    End If
    
    住院虚拟结算_渝北农医 = "个人帐户;" & dbl帐户支付 & ";0|优惠金额;" & dbl优惠金额 & ";0"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_渝北农医(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    Dim StrInput As String
    Dim str就诊流水号 As String, str结算流水号 As String
    Dim lng住院天数 As Long, lng主页ID As Long
    Dim bln正常结算 As Boolean
    Dim dbl进入统筹 As Double, dbl统筹报销 As Double, dbl现金 As Double, dbl优惠金额 As Double
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '必须先出院，才能进行结算
    If Not 医保病人已经出院(lng病人ID) Then
        Err.Raise 9000, gstrSysName, "必须先出院，才能进行结算！"
        Exit Function
    End If
    
    '取出院方式
    gstrSQL = "Select 主页ID,出院方式,住院天数 From 病案主页 Where (病人ID,主页ID) in (Select 病人ID,住院次数 From 病人信息 Where 病人ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出院方式", lng病人ID)
    lng主页ID = Nvl(rsTemp!主页ID, 1)
    lng住院天数 = Nvl(rsTemp!住院天数, 0)
    bln正常结算 = IIf(rsTemp!出院方式 = "转院", False, True)
    
    '获取病人的就诊流水号
    gstrSQL = "Select 顺序号 From 保险帐户 Where 险类=" & TYPE_渝北农医 & " And 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人的就诊流水号", lng病人ID)
    str就诊流水号 = Nvl(rsTemp!顺序号)
    
    '进行住院结算，入参：病人就诊登记的流水号|结算类别|住院的床日|经办单位|经办人|经办日期
    StrInput = str就诊流水号 & gstrSplit_渝北农医 & IIf(bln正常结算, "01", "02") & gstrSplit_渝北农医 & _
        lng住院天数 & gstrSplit_渝北农医 & gComInfo_渝北农医.医院编码 & gstrSplit_渝北农医 & _
        UserInfo.姓名 & gstrSplit_渝北农医 & Format(zlDatabase.Currentdate, mstrDateFormat) & gstrSplit_渝北农医 & _
        gstrSplit_渝北农医 & Format(zlDatabase.Currentdate, mstrDateFormat)
    Call 调用接口_准备_渝北农医(正式结算_渝北农医, StrInput)
    If Not 调用接口_渝北农医() Then Exit Function
    
    '出参：执行代码│结算流水号│这次结算医院总的金额│经过医院下浮后的总金额│合作医疗办公室承认可以参加报销的金额│实际报销的金额│病人自负的金额
    str结算流水号 = Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(1)
    dbl优惠金额 = Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(2)) - Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(3))
    dbl进入统筹 = Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(4))
    dbl统筹报销 = Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(5))
    dbl现金 = Val(Split(gstrOutput_渝北农医, gstrSplit_渝北农医)(6))
    
    '保存本次结算情况
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_渝北农医 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng主页ID & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_渝北农医.总费用 & "," & dbl现金 & "," & dbl优惠金额 & "," & dbl进入统筹 & "," & dbl统筹报销 & ",0,0," & _
        dbl统筹报销 & ",'" & str就诊流水号 & "|" & str结算流水号 & "',null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算数据")

    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "将结帐记录打上上传标志")
    
    住院结算_渝北农医 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_渝北农医(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    Dim lng冲销ID As Long
    Dim str结算流水号 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '取冲销ID
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B where A.NO=B.NO and A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!ID
    
    '取结算流水号
    gstrSQL = "Select * From 保险结算记录 Where 性质=2 And 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取结算流水号", lng结帐ID)
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "没有找到原始结算记录，无法进行住院结算冲销！", vbInformation, gstrSysName
        Exit Function
    End If
    str结算流水号 = Split(rsTemp!支付顺序号, gstrSplit_渝北农医)(1)
    
    '调用结算冲销
    Call 调用接口_准备_渝北农医(结算作废_渝北农医, str结算流水号)
    If Not 调用接口_渝北农医() Then Exit Function
    
    '保存本次结算情况
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_渝北农医 & "," & rsTemp!病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & Nvl(rsTemp!主页ID, 1) & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & "," & -1 * Nvl(rsTemp!首先自付金额, 0) & "," & -1 * Nvl(rsTemp!进入统筹金额, 0) & "," & -1 * Nvl(rsTemp!统筹报销金额, 0) & ",0,0," & _
        -1 * Nvl(rsTemp!个人帐户支付, 0) & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "门诊结算冲销")
    
    住院结算冲销_渝北农医 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Private Function IsYBPatient(ByVal lng病人ID As Long, str就诊流水号 As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '判断指定病人本次是否以医保身份就诊
    gstrSQL = " Select 1 From 病案主页 Where 险类=" & TYPE_渝北农医 & " And (病人ID,主页ID) IN " & _
              "     (Select 病人ID,住院次数 From 病人信息 Where 病人ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断指定病人本次是否以医保身份就诊", lng病人ID)
    IsYBPatient = (rsTemp.RecordCount <> 0)
    
    If IsYBPatient Then
        '取病人的就诊流水号
        gstrSQL = "Select 顺序号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人的就诊流水号", TYPE_渝北农医, lng病人ID)
        str就诊流水号 = Nvl(rsTemp!顺序号)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function CopyNewRec(ByVal SourceRec As ADODB.Recordset) As ADODB.Recordset
    Dim RecTarget As New ADODB.Recordset
    Dim intFields As Integer
    Dim intRecords As Integer
    '编制人:朱玉宝
    '编制日期:2000-11-02
    '也使用于保存
    Set RecTarget = New ADODB.Recordset
    
    With RecTarget
        If .State = 1 Then .Close
        For intFields = 0 To SourceRec.Fields.Count - 1
            .Fields.Append SourceRec.Fields(intFields).Name, adLongVarChar, 100, adFldIsNullable     '0:表示新增
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        Do While Not SourceRec.EOF
            If Nvl(SourceRec!是否上传, 0) = 0 Then
                .AddNew
                For intFields = 0 To SourceRec.Fields.Count - 1
                    .Fields(intFields) = SourceRec.Fields(intFields).Value
                Next
                .Update
            End If
            If Nvl(SourceRec!是否上传, 0) = 0 Then
                intRecords = intRecords + 1
                If intRecords = 20 Then
                    SourceRec.MoveNext
                    Exit Do
                End If
            End If
            SourceRec.MoveNext
        Loop
    End With
    
    Set CopyNewRec = RecTarget
End Function

Private Function GetSequence(ByVal StrInput As String) As String
    Dim intDO As Integer, intPos As Integer
    Dim strText As String, strSequence As String
    
    intPos = 1
    For intDO = 1 To 6
        strText = Mid(StrInput, intPos, 2)
        intPos = intPos + 2
        strSequence = strSequence & Chr(asc("0") + Val(strText))
    Next
    GetSequence = strSequence
End Function

Public Function OpenSQLServer(cnYB As ADODB.Connection, ByVal strServer As String, ByVal strUser As String, ByVal strPass As String, Optional ByVal strDatabase As String = "") As Boolean
    On Error GoTo errHand
    With cnYB
        If .State = 1 Then .Close
        .Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & strUser & ";Password=" & strPass & ";Initial Catalog=" & strDatabase & ";Data Source=" & strServer
    End With
    
    OpenSQLServer = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub WriteInfo(ByVal strInfo As String)
    Dim strFileName As String
    Dim objSystem As FileSystemObject
    Dim objStream As TextStream
    
    strFileName = "C:\YBNY_" & Format(Date, "YYYYMMdd") & ".txt"
    Set objSystem = New FileSystemObject
    If Not objSystem.FileExists(strFileName) Then Call objSystem.CreateTextFile(strFileName, False)
    Set objStream = objSystem.OpenTextFile(strFileName, ForAppending, False, TristateMixed)
    objStream.WriteLine (strInfo)
    objStream.Close
End Sub
