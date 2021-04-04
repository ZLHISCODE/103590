Attribute VB_Name = "mdl乐山"
Option Explicit
Public Declare Sub LS_ErrMessage Lib "SIHisInterface.dll" Alias "GetErrorMessage" (ErrorMsg As TStringOfChar)
Public Declare Function LS_UserLogin Lib "SIHisInterface.dll" Alias "UserLogin" (UserCode As TStringOfChar, PWD As TStringOfChar) As Byte
Public Declare Function LS_ChangePwd Lib "SIHisInterface.dll" Alias "ChangeUserPwd" (OldPwd As TStringOfChar, NewPWD As TStringOfChar) As Byte
Public Declare Sub LS_UserLogout Lib "SIHisInterface.dll" Alias "UserLogout" ()
Public Declare Function LS_ConnectServer Lib "SIHisInterface.dll" Alias "ConnectServer" (ServerName As TStringOfChar) As Byte
Public Declare Sub LS_DisConnectServer Lib "SIHisInterface.dll" Alias "DisConnectServer" ()

'获取参保人信息
Public Declare Function LS_GetPersonInfo Lib "SIHisInterface.dll" Alias "GetPersonInfo" (PInfo As 身份信息) As Byte
'入院登记
Public Declare Function LS_InHospitalRegister Lib "SIHisInterface.dll" Alias "InBedRegster" (InBedRegInfo As 住院信息) As Byte
'获取入院登记信息
Public Declare Function LS_GetInHospitalRegInfo Lib "SIHisInterface.dll" Alias "GetInBedRegInfo" (InBedRegID As TStringOfChar) As Byte
'录入药品费用
Public Declare Function LS_AddDrug Lib "SIHisInterface.dll" Alias "AddDrug" (DrugInfo As 药品信息) As Byte
'录入诊疗费用
Public Declare Function LS_AddDiag Lib "SIHisInterface.dll" Alias "AddDiag" (DiagInfo As 诊疗信息) As Byte
'录入服务设施费用
Public Declare Function LS_AddService Lib "SIHisInterface.dll" Alias "AddServiceItem" (ServiceItemInfo As 服务设施信息) As Byte
'保存费用明细
Public Declare Function LS_SaveDetail Lib "SIHisInterface.dll" Alias "InBedRegApplyUpdates" (InBedRegID As TStringOfChar) As Byte
'住院费用预结算
Public Declare Function LS_PreBalance Lib "SIHisInterface.dll" Alias "NewInBedBill" (InBedBillInfo As 住院结算信息) As Byte
'住院费用结算
Public Declare Function LS_Balance Lib "SIHisInterface.dll" Alias "SaveInBedBill" (InBedBillInfo As 住院结算信息) As Byte

'----增加门诊业务----
Public Declare Function LS_ExamBill Lib "SIHisInterface.dll" Alias "NewExamBill" (TexamBillInfo As 门诊结算单) As Byte
'录入门诊药品费用
Public Declare Function LS_ExamAddDrug Lib "SIHisInterface.dll" Alias "AddExamDrug" (DrugInfo As 药品信息) As Byte
'录入门诊诊疗费用
Public Declare Function LS_ExamAddDiag Lib "SIHisInterface.dll" Alias "AddExamDiag" (DiagInfo As 诊疗信息) As Byte
'录入门诊服务设施费用
Public Declare Function LS_ExamAddServiceItem Lib "SIHisInterface.dll" Alias "AddExamServiceItem" (ServiceItemInfo As 服务设施信息) As Byte
'门诊预结算
Public Declare Function LS_ExamPreBalance Lib "SIHisInterface.dll" Alias "ExamBillReCalculate" (TexamBillInfo As 门诊结算单) As Byte
'门诊预结算
Public Declare Function LS_ExamBalance Lib "SIHisInterface.dll" Alias "SaveExamBill" (TexamBillInfo As 门诊结算单) As Byte
'--------------------

'全局变量区
Private Const mstr出院结帐 As String = "出院结帐"
Private Const mstr中途暂结帐 As String = "中途暂结帐"
Private Const mstr转院结帐 As String = "转院结帐"

'病人相关信息常量
Private Const 入院科室编号 = 0
Private Const 入院科室名称 = 1
Private Const 入院病区编号 = 2
Private Const 入院病区名称 = 3
Private Const 入院病床编号 = 4
Private Const 出院科室名称 = 5
Private Const 住院医师 = 6
Private Const 住院号 = 7
Private Const 入院诊断 = 8
Private Const 出院诊断 = 9
Private Const 出院日期 = 10
Private Const 出院方式 = 11

Public Type TStringOfChar
    Data As String * 100
End Type
Public Type 身份信息                   'TPersonInfo
    '以下数据为返回数据
    PSN_ID              As Long      '医疗参保ID号
    PSN_No              As Long      '参保人编码
    PSN_NAME            As String * 100 '参保人姓名
    Sex                 As String * 100 '性别
    IDCARD              As String * 100 '身份证号码
    PSN_STS             As String * 100 '参保人状态
    PSN_TYP             As String * 100 '人员类别
    UNIT_CODE           As String * 100 '单位编码
    UNIT_NAME           As String * 100 '单位名称
    OFFICAL_TYP         As String * 100 '公务员类别
    HAI_TYP             As String * 100 '补充医保名称
    ACCT_STS            As String * 100 '医保账户状态
    HI_ACCT_PWD         As String * 100 '医保帐户口令
    SILL_PAY_AMT_TOTAL  As Single       '年内进入门诊特殊疾病支付金额
    SILL_YR_FUND_AMT    As Single       '年内门诊统筹基金支付金额
    YR_FUND_AMT         As Single       '年内统筹基金支付金额
    HAI_YR_HIGH_AMT     As Single       '年内补充高额支付金额
    HAI_YR_INBED_AMT    As Single       '年内补充住院补助支付金额
    GZ_CUR_AMT          As Single       '个人账户余额
    YR_INBED_CNT        As Long      '年内住院次数
    CARD_NO             As String * 100 'IC卡号
End Type
Private Type 住院信息                   'TInBedRegInfo
    PSN_ID              As Long      '医疗参保人ID号
    INBED_SILL_ID       As Long      '住院特殊病种ID（保留）
    INBED_NO            As String * 100 '住院号
    INBED_EXAM          As String * 100 '入院诊断
    INBED_EXAM_ICD10_NO As String * 100 '入院诊断ICD10编码
    INBED_DEPT          As String * 100 '入院科室
    '以下数据为返回数据
    INBED_REG_ID        As String * 100 '住院登记ID
    INBED_DT            As String * 100 '入院时间，录入数据
End Type
Private Type 药品信息               'TDrugInfo
    INBED_REG_ID    As String * 100 '住院登记ID
    RECEIPT_DT      As String * 100 '收费时间
    DRUG_CATALOG_ID As String * 100 '药品代码参数ID
    DRUG_INFO       As String * 100 '药品信息
    UNIT_PRC        As Single       '单价
    SRVC_CNT        As Single       '数量
    COST_PRC        As Single       '成本单价
    DRUG_TYP        As String * 100 '药物剂型
    DRUG_SPEC       As String * 100 '药物规格
    PRODUCE_FACTORY As String * 100 '生产厂家
    '以下数据为返回数据
    FEE_ITEM_TYP    As String * 100 '费用项目种类
    FEE_TYP         As String * 100 '费用种类
    PART_PUB_AMT    As Single       '部分公费金额
    PART_SELF_AMT   As Single       '部分自费金额
    PUB_PAY_AMT     As Single       '公费金额
    SELF_PAY_AMT    As Single       '自费金额
    SELF_PAY_PCT    As Single       '自费比例
    MAX_RETAIL_PRC  As Single       '最高零售价
    DRUG_SPC_FLAG   As Long         '特殊用药标志
End Type
Private Type 诊疗信息               'TDiagInfo
    INBED_REG_ID    As String * 100 '住院登记ID
    RECEIPT_DT      As String * 100 '收费时间
    DIAG_CATALOG_ID As String * 100 '诊疗项目代码参数ID
    DIAG_ITEM_NAME  As String * 100 '诊疗项目名称
    UNIT_PRC        As Single       '单价
    SRVC_CNT        As Single       '数量
    '以下数据为返回数据
    FEE_ITEM_TYP    As String * 100 '费用项目种类
    FEE_TYP         As String * 100 '费用种类
    PART_PUB_AMT    As Single       '部分公费金额
    PART_SELF_AMT   As Single       '部分自费金额
    PUB_PAY_AMT     As Single       '公费金额
    SELF_PAY_AMT    As Single       '自费金额
    SELF_PAY_PCT    As Single       '自费比例
    MAX_RETAIL_PRC  As Single       '最高零售价
End Type
Private Type 服务设施信息           'TServiceItemInfo
    INBED_REG_ID    As String * 100 '住院登记ID
    RECEIPT_DT      As String * 100 '收费时间
    SRVC_ITEM_ID    As String * 100 '基本医疗保险服务设施标准
    SRVC_NAME       As String * 100 '服务设施名称
    UNIT_PRC        As Single       '单价
    SRVC_CNT        As Single       '数量
    '以下数据为返回数据
    FEE_ITEM_TYP    As String * 100 '费用项目种类
    FEE_TYP         As String * 100 '费用种类
    PART_PUB_AMT    As Single       '部分公费金额
    PART_SELF_AMT   As Single       '部分自费金额
    PUB_PAY_AMT     As Single       '公费金额
    SELF_PAY_AMT    As Single       '自费金额
    SELF_PAY_PCT    As Single       '自费比例
    MAX_RETAIL_PRC  As Single       '最高零售价
End Type
Private Type 住院结算信息                   'TInBedBillInfo
    INBED_REG_ID        As String * 100     '住院登记ID
    EXAM_TYP            As String * 100     '就诊类别
    INBED_STL_TYP       As String * 100     '住院结帐方式
    OUTBED_EXAM         As String * 100     '出院诊断
    OUTBED_EXAM_ICD10_NO As String * 100    '出院诊断ICD10编码
    OUTBED_DEPT         As String * 100     '出院科室
    ILL_TRS_STS         As String * 100     '疾病转归(治愈、死亡…)
    INBED_DOCTOR        As String * 100     '管床医生
    OUTBED_DT           As String * 100     '出院时间
    '以下数据为返回数据
    INBED_DAY_CNT       As Long          '住院天数
    FEE_STL_LOC         As String * 100     '费用结算地点
    EXAM_ADDR           As String * 100     '就诊地点
    INBED_STL_BILL_ID   As String * 100     '住院结帐单id
    INBED_STL_BILL_NO   As String * 100     '住院结帐单号
    PART_PUB_AMT        As Single           '部分公费金额
    PART_SELF_AMT       As Single           '部分自费金额
    PUB_PAY_AMT         As Single           '公费金额
    SELF_PAY_AMT        As Single           '自费金额
    INBED_FUND_AMT      As Single           '住院统筹支付金额
    INBED_ACCT_AMT      As Single           '住院个账支付金额
    CASH_PAY_AMT        As Single           '现金支付金额
    HAI_INBED_SBS_AMT   As Single           '补充住院补助支付金额
    HAI_INBED_AMT       As Single           '补充住院支付金额
    HAI_INBED_REPAY_AMT As Single           '补充住院再次支付金额
    HAI_INBED_HIGH_AMT  As Single           '补充住院高额支付金额
    OFFICAL_HIGH_AMT    As Single           '公务员高额补助支付金额
    OFFICAL_INBED_AMT   As Single           '公务员住院补助支付金额
    OFFICAL_ACCT_AMT    As Single           '公务员个帐补助支付金额
End Type
'----增加门诊业务----
Private Type 门诊结算单
    PSN_ID           As Long                '医疗参保人ID号
    EXAM_TYP         As String * 100        '就诊类别
    EXAM_DEPT        As String * 100        '门诊科室
    EXAM_DOCTOR      As String * 100        '门诊医生
    '以下数据为返回数据
    FEE_STL_LOC      As String * 100        '费用结算地点
    EXAM_ADDR        As String * 100        '就诊地点
    EXAM_STL_BILL_ID As String * 100        '门诊结帐单id
    EXAM_STL_BILL_NO As String * 100        '门诊结帐单号
    PART_PUB_AMT     As Single              '部分公费金额
    PART_SELF_AMT    As Single              '部分自费金额
    PUB_PAY_AMT      As Single              '公费金额
    SELF_PAY_AMT     As Single              '自费金额
    EXAM_ACCT_AMT    As Single              '门诊个账支付金额
    CASH_PAY_AMT     As Single              '门诊现金支付金额
End Type
'--------------------

Private Type 结算信息
    顺序号 As TStringOfChar
    总费用 As Currency
    现金 As Currency
    个人帐户 As Currency
    医保基金 As Currency
    补充基金 As Currency
    公务员基金 As Currency
    医保总费用 As Currency
End Type
Public gPersonInfo_乐山 As 身份信息
Public gInBedRegInfo_乐山 As 住院信息
Public gDrugInfo_乐山 As 药品信息
Public gDiagInfo_乐山 As 诊疗信息
Public gServiceItemInfo_乐山 As 服务设施信息
Public gInBedBillInfo_乐山 As 住院结算信息
Private gtypBalance As 结算信息
Private gExamBill As 门诊结算单

Private glngInterface_乐山 As Long
Private gstrErrMsg_乐山 As TStringOfChar          '错误信息
Public gbytReturn_乐山 As Byte                '0-正常;非零值代表错误号

Public Function 医保初始化_乐山() As Boolean
    Dim strServer As TStringOfChar
    On Error GoTo errHand
    
    If glngInterface_乐山 <> 0 Then 医保初始化_乐山 = True: Exit Function
    strServer = GetServerInfo
    If strServer.Data = "" Then Exit Function
    
    '连接服务器
    gbytReturn_乐山 = LS_ConnectServer(strServer)
    If GetErrInfo_乐山 Then Exit Function
    
    '登录中心(失败则断开连接并退出)
    If Not frm登录中心.LoginCenter(TYPE_乐山, True) Then
        Call 医保终止_乐山
        Exit Function
    End If
    glngInterface_乐山 = 1
    
    医保初始化_乐山 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
  Resume
    End If
End Function

Public Function 医保终止_乐山() As Boolean
    On Error Resume Next
    If glngInterface_乐山 = 0 Then
        医保终止_乐山 = True
        Exit Function
    End If

    '操作员退出
    Call LS_UserLogout
    '连接服务器
    Call LS_DisConnectServer
    glngInterface_乐山 = 0
    
    医保终止_乐山 = True
End Function

Public Function 医保设置_乐山() As Boolean
    With frmSet乐山
        医保设置_乐山 = .ShowME
    End With
End Function

Public Function GetErrInfo_乐山() As Boolean
    If gbytReturn_乐山 = 1 Then Exit Function
    Call LS_ErrMessage(gstrErrMsg_乐山)
    MsgBox gstrErrMsg_乐山.Data, vbInformation, gstrSysName
    GetErrInfo_乐山 = True
End Function

Private Function GetServerInfo() As TStringOfChar
    '获取服务器地址
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '获取服务器地址、端口及入口名称('服务器地址','服务器端口号','服务器入口程序')
    gstrSQL = " Select 参数名,参数值 From 保险参数" & _
              " Where 险类=[1] And 参数名 = '服务器地址'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取服务器名称或IP地址", TYPE_乐山)
    
    With rsTemp
        If .RecordCount = 0 Then Exit Function
        GetServerInfo.Data = Nvl(!参数值)
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 个人余额_乐山(ByVal str医保号 As String) As Currency
    '功能: 直接读出卡内金额
    '参数: 是否读卡
    '返回: 返回个人帐户余额
    Dim rsAccount As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = " Select Nvl(帐户余额,0) 帐户余额 From 保险帐户 " & _
              " Where 险类=[1] And 医保号=[2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "返回个人帐户余额", TYPE_乐山, str医保号)
    
    个人余额_乐山 = rsAccount!帐户余额
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 门诊虚拟结算_乐山(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    'cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保,开单人
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    On Error GoTo errHand
    Dim intType As Integer
    Dim str开单时间 As String
    Dim rs大类 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim dbl单价 As Double
    
    If Nvl(rs明细!开单人) = "" Then
        MsgBox "开单医生不能为空！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '由于未传递开单科室,只有根据开单人,提取开单科室(只能取临床科室)
    Call DebugTool("提取开单医生所在科室")
    gstrSQL = "SELECT C.名称 AS 开单科室 " & _
             " FROM 部门人员 A,部门性质说明 B,部门表 C " & _
             " WHERE A.人员ID= " & _
             "     (SELECT ID FROM 人员表 WHERE 姓名=[1]) " & _
             " AND A.部门ID=B.部门ID AND A.部门ID=C.ID AND B.工作性质='临床' AND 服务对象 IN (1,3) " & _
             " AND ROWNUM<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取开单科室", CStr(rs明细!开单人))
    If rsTemp.EOF Then
        MsgBox "该医生不属于任何临床科室！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '创建新的门诊结算单
    Call DebugTool("创建新的门诊结算单")
    With gExamBill
        .PSN_ID = gPersonInfo_乐山.PSN_ID       '医疗参保人ID
        .EXAM_TYP = ""                          '就诊类别
        .EXAM_DEPT = rsTemp!开单科室            '门诊科室
        .EXAM_DOCTOR = Nvl(rs明细!开单人)       '门诊医生
    End With
    gbytReturn_乐山 = LS_ExamBill(gExamBill)
    If GetErrInfo_乐山 Then Exit Function
    
    Call DebugTool("获取开单时间及病人基本信息")
    str开单时间 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    Call 获取病人基本信息(rs明细!病人ID)
    
    '提取保险大类
    Call DebugTool("提取保险大类")
    gstrSQL = "Select ID,名称 From 保险支付大类 Where 险类=[1]"
    Set rs大类 = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保大类", TYPE_乐山)
    
    '上传明细
    Call DebugTool("准备明细上传")
    gtypBalance.总费用 = 0
    With rs明细
        Do While Not .EOF
            gtypBalance.总费用 = gtypBalance.总费用 + Nvl(!实收金额, 0)
            If rs明细!实收金额 = 0 Then
                .MoveNext
                If .EOF Then
                    Exit Do
                End If
            End If
            '上传明细
            intType = 1
            rs大类.Filter = "ID=" & !保险支付大类ID
            If rs大类.RecordCount <> 0 Then
                If rs大类!名称 = "诊疗" Then intType = 2
                If rs大类!名称 = "服务" Then intType = 3
            End If
            rs大类.Filter = 0
            
            Select Case intType
            Case 1
                Call DebugTool("取药品信息")
                gstrSQL = "Select A.规格,A.产地,B.名称 AS 剂型,D.项目编码 AS 医保项目编码,E.名称 AS 细目名称" & _
                         " From 药品目录 A,药品剂型 B,药品信息 C,保险支付项目 D,收费细目 E " & _
                         " Where A.药名ID=C.药名ID And C.剂型=B.编码 And A.药品ID=" & !收费细目ID & _
                         " And D.险类=[1] And E.ID=D.收费细目ID And D.收费细目ID=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取药品信息", TYPE_乐山, CLng(!收费细目ID))
                
                With gDrugInfo_乐山
                    .INBED_REG_ID = ""
                    .RECEIPT_DT = str开单时间
                    .DRUG_CATALOG_ID = rsTemp!医保项目编码
                    .DRUG_INFO = rsTemp!细目名称
                    dbl单价 = rs明细!实收金额 / rs明细!数量
                    '曾明春:2005-07-06 如果单价精度超过2位小数，则数量传1，单价传实收金额。
                    If Round(dbl单价 * 100) <> dbl单价 * 100 Then
                        '曾明春:2005-06-02修改，金额取绝对值,避免单价出现负数的情况
                        .UNIT_PRC = Format(Abs(rs明细!实收金额), "#####0.00;-#####0.00;0;")
                        If rs明细!实收金额 <= 0 Then
                          .SRVC_CNT = -1
                        Else
                          .SRVC_CNT = 1
                        End If
                    Else
                        .UNIT_PRC = Format(rs明细!实收金额 / rs明细!数量, "#####0.00;-#####0.00;0;")
                        .SRVC_CNT = rs明细!数量
                    End If
                    .COST_PRC = 0
                    .DRUG_TYP = Nvl(rsTemp!剂型)
                    .DRUG_SPEC = Nvl(rsTemp!规格)
                    .PRODUCE_FACTORY = Nvl(rsTemp!产地)
                    .DRUG_SPC_FLAG = 0
                End With
            Case 2
                Call DebugTool("取诊疗信息")
                gstrSQL = "Select D.项目编码 AS 医保项目编码,E.名称 AS 细目名称" & _
                         " From 保险支付项目 D,收费细目 E " & _
                         " Where D.险类=[1] And E.ID=D.收费细目ID And D.收费细目ID=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取药品信息", TYPE_乐山, CLng(!收费细目ID))
                
                With gDiagInfo_乐山
                    .INBED_REG_ID = ""
                    .RECEIPT_DT = str开单时间
                    .DIAG_CATALOG_ID = rsTemp!医保项目编码
                    .DIAG_ITEM_NAME = rsTemp!细目名称
                    dbl单价 = rs明细!实收金额 / rs明细!数量
                    '曾明春:2006-11-06 如果单价精度超过2位小数，则数量传1，单价传实收金额。
                    If Round(dbl单价 * 100) <> dbl单价 * 100 Then
                        '曾明春:2006-11-06修改，金额取绝对值,避免单价出现负数的情况
                        .UNIT_PRC = Format(Abs(rs明细!实收金额), "#####0.00;-#####0.00;0;")
                        If rs明细!实收金额 <= 0 Then
                          .SRVC_CNT = -1
                        Else
                          .SRVC_CNT = 1
                        End If
                    Else
                        .UNIT_PRC = Format(rs明细!实收金额 / rs明细!数量, "#####0.00;-#####0.00;0;")
                        .SRVC_CNT = rs明细!数量
                    End If
                End With
            Case 3
                Call DebugTool("取服务设施信息")
                gstrSQL = "Select D.项目编码 AS 医保项目编码,E.名称 AS 细目名称" & _
                         " From 保险支付项目 D,收费细目 E " & _
                         " Where D.险类=[1] And E.ID=D.收费细目ID And D.收费细目ID=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取药品信息", TYPE_乐山, CLng(!收费细目ID))
                
                With gServiceItemInfo_乐山
                    .INBED_REG_ID = ""
                    .RECEIPT_DT = str开单时间
                    .SRVC_ITEM_ID = rsTemp!医保项目编码
                    .SRVC_NAME = rsTemp!细目名称
                    dbl单价 = rs明细!实收金额 / rs明细!数量
                    '曾明春:2006-11-06 如果单价精度超过2位小数，则数量传1，单价传实收金额。
                    If Round(dbl单价 * 100) <> dbl单价 * 100 Then
                        '曾明春:2006-11-06修改，金额取绝对值,避免单价出现负数的情况
                        .UNIT_PRC = Format(Abs(rs明细!实收金额), "#####0.00;-#####0.00;0;")
                        If rs明细!实收金额 <= 0 Then
                          .SRVC_CNT = -1
                        Else
                          .SRVC_CNT = 1
                        End If
                    Else
                        .UNIT_PRC = Format(rs明细!实收金额 / rs明细!数量, "#####0.00;-#####0.00;0;")
                        .SRVC_CNT = rs明细!数量
                    End If
                End With
            End Select
            
            Call DebugTool("上传明细")
            If Not UploadDetail(intType, False) Then Exit Function
            .MoveNext
        Loop
    End With
    
    '预结算
    '个人帐户支付额:EXAM_ACCT_AMT,不允许修改,门诊仅支持个人帐户,不支持医保基金
    Call DebugTool("门诊预结算")
    gbytReturn_乐山 = LS_ExamPreBalance(gExamBill)
    If GetErrInfo_乐山 Then Exit Function
    
    '基金数据赋值
    Call DebugTool("获取各项基金支付额")
    With gtypBalance
        .个人帐户 = gExamBill.EXAM_ACCT_AMT
        .补充基金 = 0
        .公务员基金 = 0
        .医保基金 = 0
        .现金 = .总费用 - .个人帐户
    End With
    
    str结算方式 = "个人帐户;" & gtypBalance.个人帐户 & ";0"
    门诊虚拟结算_乐山 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 门诊结算_乐山(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    'cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '注意：接口规定，门诊明细需结算后上传；住院明细需预结算时上传，如果卡内金额不足，可以使用圈存接口，即将卡外的钱，调到卡内，以增加卡内金额
    '卡内余额需要通过卡操作函数读取，可圈存金额是接口返回，需要修改
    On Error GoTo errHand
    Dim lng病人ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim datCurr As Date
    
    '门诊结算
    Call DebugTool("正式结算")
    gbytReturn_乐山 = LS_ExamBalance(gExamBill)
    If GetErrInfo_乐山 Then Exit Function
    
    Call DebugTool("提取病人ID")
    gstrSQL = "Select 病人ID,登记时间 From 门诊费用记录 Where 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人ID", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    datCurr = rsTemp!登记时间
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur起付线累计 As Currency, cur本次起付线 As Currency, cur统筹限额 As Currency
    
    '帐户年度信息
    Call Get帐户信息(TYPE_乐山, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计, cur本次起付线, cur起付线累计, cur统筹限额)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_乐山 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur本次起付线 & "," & cur起付线累计 & "," & cur统筹限额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    Call DebugTool("保存保险结算记录")
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_乐山 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gtypBalance.总费用 & "," & gtypBalance.现金 & "," & 0 & "," & gtypBalance.医保基金 & "," & gtypBalance.医保基金 & "," & _
        gtypBalance.补充基金 & "," & gtypBalance.公务员基金 & "," & gtypBalance.个人帐户 & ",'" & TrimTsChar(gtypBalance.顺序号.Data) & "',null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊结算数据")
    
    门诊结算_乐山 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_乐山(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    'cur个人帐户   从个人帐户中支出的金额
    On Error GoTo errHand
    
    Err.Raise 9000, gstrSysName, "乐山医保不支持门诊退费，请与医保中心联系!"
    门诊结算冲销_乐山 = False
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 入院登记_乐山(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    Dim str顺序号 As String, strID As String
    Dim arrPatient
    
    On Error GoTo errHand
    
    arrPatient = Split(获取病人相关信息(lng病人ID, lng主页ID), "||")
    '写传入参数
    With gInBedRegInfo_乐山
        .PSN_ID = gPersonInfo_乐山.PSN_ID                           '住院参保ID号
        .INBED_SILL_ID = 0                                          '住院特殊病种ID（保留）
        .INBED_NO = arrPatient(住院号)                              '住院号
        .INBED_EXAM = Split(arrPatient(入院诊断), "|")(0)           '入院诊断
        .INBED_EXAM_ICD10_NO = Split(arrPatient(入院诊断), "|")(1)  '入院诊断ICD10编码
        .INBED_DEPT = arrPatient(入院科室名称)                          '入院科室
    End With
    
    '调用入院登记接口
    gbytReturn_乐山 = LS_InHospitalRegister(gInBedRegInfo_乐山)
    If GetErrInfo_乐山 Then Exit Function
    
    '更新个人帐户中的信息
    str顺序号 = TrimTsChar(gInBedRegInfo_乐山.INBED_REG_ID)
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_乐山 & ",'顺序号','''" & str顺序号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存入院业务序列号")
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_乐山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")

    入院登记_乐山 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_乐山(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '功能：将出院信息发送医保前置服务器确认
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    '不允许撤销入院
    On Error GoTo errHand
    
    MsgBox "不支持出院登记撤销，请与医保接口商联系！", vbInformation, gstrSysName
    入院登记撤销_乐山 = False
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_乐山(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHand
    '功能：将出院信息发送医保前置服务器确认
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false

    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_乐山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")

    出院登记_乐山 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记撤销_乐山(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rs费用          As ADODB.Recordset
    On Error GoTo errHand
'    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_乐山 & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销出院登记")
    MsgBox "由于该地区不支持入院登记撤销，本次操作只针对本地数据撤销，取消中心登记请到医保软件取消或联系医保接口商！", vbInformation, gstrSysName
    gstrSQL = "ZL_病案主页_撤消医保入院(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "撤消医保入院")
    gstrSQL = "select id from 住院费用记录 where 病人id = [1] and 主页id = [2]"
    Set rs费用 = zlDatabase.OpenSQLRecord(gstrSQL, "查出所有该次住院费用记录", lng病人ID, lng主页ID)
    If Not rs费用.EOF Then
        Do While Not rs费用.EOF
        '考虚可能再次补充成保险病人时，费用需重新上传,该处置为0
            gstrSQL = "ZL_病人费用记录_更新医保(" & rs费用!ID & ",null,null,null,null" & ",0,null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病人费用记录中所有记录上传标志为0")
            rs费用.MoveNext
        Loop
    End If
    出院登记撤销_乐山 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 住院虚拟结算_乐山(rsExse As Recordset, ByVal lng病人ID As Long) As String
    Dim lng主页ID As Long
    Dim bln出院结算 As Boolean
    Dim str记录性质 As String, str记录状态 As String, strNO As String
    Dim arrPatient
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
    
    '读取主页ID
    gstrSQL = "Select 住院次数 主页ID From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取主页ID", lng病人ID)
    lng主页ID = rsTemp!主页ID
    
    '获取总费用
    gtypBalance.总费用 = 0
    With rsExse
        Do While Not .EOF
            gtypBalance.总费用 = gtypBalance.总费用 + Nvl(!金额, 0)
            '上传明细
            If !主页ID = lng主页ID And Nvl(!是否上传, 0) = 0 And (strNO <> !NO Or str记录性质 <> !记录性质 Or str记录状态 <> !记录状态) Then
                strNO = !NO
                str记录性质 = !记录性质
                str记录状态 = !记录状态
                If Not 上传处方_乐山(str记录性质, str记录状态, strNO) Then Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    Call 获取病人基本信息(lng病人ID)
    arrPatient = Split(获取病人相关信息(lng病人ID, lng主页ID), "||")
    bln出院结算 = 医保病人已经出院(lng病人ID)
    
    '写传入参数
    '曾明春(2006-05-15):出院方式中增加治愈、好转、未愈，将转院、死亡、治愈、好转、未愈作为疾病转归方式；如果转院方式为正常，则疾病转归默认为好转。
    With gInBedBillInfo_乐山
        .INBED_REG_ID = gtypBalance.顺序号.Data
        .EXAM_TYP = ""
        .INBED_STL_TYP = IIf(bln出院结算, IIf(arrPatient(出院方式) = "转院", mstr转院结帐, mstr出院结帐), mstr中途暂结帐)
        .OUTBED_EXAM = Split(arrPatient(出院诊断), "|")(0)
        .OUTBED_EXAM_ICD10_NO = Split(arrPatient(出院诊断), "|")(1)
        .OUTBED_DEPT = arrPatient(出院科室名称)
        .ILL_TRS_STS = IIf(bln出院结算, IIf(arrPatient(出院方式) = "正常", "治愈", arrPatient(出院方式)), "未愈")
        .INBED_DOCTOR = arrPatient(住院医师)
        .OUTBED_DT = IIf(bln出院结算, arrPatient(出院日期), "")
    End With
    gbytReturn_乐山 = LS_PreBalance(gInBedBillInfo_乐山)
    If GetErrInfo_乐山 Then Exit Function

    Call Get结算信息
    
    '提示参保病人的住院相关信息
    If Format(gtypBalance.总费用, "#0.00") <> Format(gtypBalance.医保总费用, "#0.00") Then
        MsgBox "该参保病人的医保总费用：￥" & Format(gtypBalance.医保总费用, "#0.00") & "元     " & vbCrLf & _
               "      与内部系统总费用：￥" & Format(gtypBalance.总费用, "#0.00") & "元     " & vbCrLf & _
               " 不一致.请联系管理员检查后再结帐!", vbInformation, gstrSysName
    End If
           
    住院虚拟结算_乐山 = "个人帐户;" & gtypBalance.个人帐户 & ";0"
    住院虚拟结算_乐山 = 住院虚拟结算_乐山 & "|医保基金;" & gtypBalance.医保基金 & ";0"
    住院虚拟结算_乐山 = 住院虚拟结算_乐山 & "|补充基金;" & gtypBalance.补充基金 & ";0"
    住院虚拟结算_乐山 = 住院虚拟结算_乐山 & "|公务员基金;" & gtypBalance.公务员基金 & ";0"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_乐山(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    Dim cur帐户支付 As Currency
    Dim rsTemp As New ADODB.Recordset
    '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
  '虚拟结算（返回的数据减去历次结算数据，就等于本次的真实结算数据）
    On Error GoTo errHand
    Call 获取病人基本信息(lng病人ID)
    
    '读取本次个人帐户支付额
    gstrSQL = "Select Nvl(A.冲预交,0) 个人帐户 " & _
        " From 病人预交记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And B.险类=[2]" & _
        " And A.结算方式 in ('个人帐户') And A.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取本次个人帐户支付额", lng结帐ID, TYPE_乐山)
    cur帐户支付 = 0
    If Not rsTemp.EOF Then
        cur帐户支付 = rsTemp!个人帐户
    End If
    
    '直接调用结算接口，因为虚拟结算已经填写了入口参数
    gbytReturn_乐山 = LS_Balance(gInBedBillInfo_乐山)
    If GetErrInfo_乐山 Then Exit Function
    
    Call Get结算信息(cur帐户支付)
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "将结帐记录打上上传标志")
    
    '填写保险结算记录
    '（大病自付金额=补充基金;超限自付金额=公务员基金）
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_乐山 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gtypBalance.总费用 & "," & gtypBalance.现金 & "," & 0 & "," & gtypBalance.医保基金 & "," & gtypBalance.医保基金 & "," & _
        gtypBalance.补充基金 & "," & gtypBalance.公务员基金 & "," & cur帐户支付 & ",'" & TrimTsChar(gtypBalance.顺序号.Data) & "',null,null,'" & TrimTsChar(gInBedBillInfo_乐山.INBED_STL_BILL_NO) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算数据")
    住院结算_乐山 = True
    
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_乐山(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    Err.Raise 9000, gstrSysName, "不支持住院结算冲销，请到医保中心办理！ "
    住院结算冲销_乐山 = False
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 身份标识_乐山(Optional bytType As Byte, Optional lng病人ID As Long) As String
'    功能：识别指定人员是否为参保病人，返回病人的信息
'    参数：bytType-识别类型，0-门诊，1-住院
'返回:     空或信息串
'    注意：1)主要利用接口的身份识别交易；
'    2)如果识别错误，在此函数内直接提示错误信息；
'    3)识别正确，而个人信息缺少某项，必须以空格填充；
    '仅支持住院
    身份标识_乐山 = frmIdentify乐山.GetPatient(bytType, lng病人ID)
End Function

Private Function 获取病人相关信息(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    Dim str入院科室编号 As String, str入院科室名称 As String, str入院病区编号 As String
    Dim str入院病区名称 As String, str入院病床编号 As String, str出院科室名称 As String
    Dim str住院医师 As String, str住院号 As String, str入院诊断 As String
    Dim str出院诊断 As String, str出院日期 As String, str出院方式 As String
    Dim rsTemp As New ADODB.Recordset
'    读取病人相关信息 (本年住院次数||入院科室编号||入院科室名称||入院病区编号||入院病区名称||入院病床编号||住院号||入院诊断||出院诊断)
    
'    读取入院相关信息
    gstrSQL = "select C.编码 入院科室编号,C.名称 入院科室名称,B.编码 入院病区编号,B.名称 入院病区名称, " & _
             " A.入院病床 入院病床编号,D.名称 出院科室名称,F.床位类型,E.住院号 住院号,A.住院医师,to_char(A.出院日期,'yyyy-MM-dd') 出院日期,A.出院方式 " & _
             " from 病案主页 A,部门表 B,部门表 C,部门表 D,病人信息 E, " & _
             " (Select D.名称 床位类型,F.床号,F.科室ID,F.病区ID  From 床位等级 D ,床位状况记录 F Where F.等级ID=D.序号) F " & _
             " Where A.入院病区ID=B.ID(+) And A.入院科室ID=C.ID(+) And A.出院科室ID=D.ID(+) And A.病人ID=E.病人ID ANd A.病人ID=[1] And A.主页ID=[2]" & _
             " And A.入院病床=F.床号(+) And F.科室ID(+)=A.入院科室ID And F.病区ID(+)=A.入院病区ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院相关信息", lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
        str入院科室编号 = Nvl(rsTemp!入院科室编号)
        str入院科室名称 = Nvl(rsTemp!入院科室名称)
        str入院病区编号 = Nvl(rsTemp!入院病区编号)
        str入院病区名称 = Nvl(rsTemp!入院病区名称)
        str入院病床编号 = Nvl(rsTemp!入院病床编号)
        str出院科室名称 = Nvl(rsTemp!出院科室名称)
        str住院医师 = Nvl(rsTemp!住院医师)
        str出院日期 = Nvl(rsTemp!出院日期)
        str出院方式 = Nvl(rsTemp!出院方式)
        str住院号 = Nvl(rsTemp!住院号)
    End If
    
'    读取入出院诊断（诊断|疾病编码）
    str入院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, True, False, True)
    str出院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, False, False, True)
    获取病人相关信息 = str入院科室编号 & "||" & str入院科室名称 & "||" & _
                    str入院病区编号 & "||" & str入院病区名称 & "||" & str入院病床编号 & "||" & _
                    str出院科室名称 & "||" & str住院医师 & "||" & str住院号 & "||" & str入院诊断 & _
                    "||" & str出院诊断 & "||" & str出院日期 & "||" & str出院方式
End Function

Private Sub 获取病人基本信息(ByVal lng病人ID As Long)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select 顺序号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人的住院流水号", lng病人ID, TYPE_乐山)
    
    gtypBalance.顺序号.Data = Nvl(rsTemp!顺序号)
End Sub

Private Function 是否医保病人(ByVal lng病人ID As Long) As Boolean
    Dim rsInsure As New ADODB.Recordset
    
    '检查本次是否以医保身份入院
    gstrSQL = "Select Count(*) Records From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.病人ID=[1] And A.主页ID=B.住院次数 And A.险类=[2]"
    Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否医保病人", lng病人ID, TYPE_乐山)
    是否医保病人 = (rsInsure!Records = 1)
End Function

Private Sub Get结算信息(Optional ByVal cur帐户支付 As Currency = 0)
    '根据预结算或结算返回的值，显示结算信息（由于个人帐户是接口返回的，估计不允许修改）
    With gtypBalance
'        PART_PUB_AMT      : Single;         //部分公费金额
'        PART_SELF_AMT     : Single;         //部分自费金额
'        PUB_PAY_AMT       : Single;         //公费金额
'        SELF_PAY_AMT      : Single;         //自费金额

'        INBED_FUND_AMT      As Single           '住院统筹支付金额
'        INBED_ACCT_AMT      As Single           '住院个账支付金额
'        CASH_PAY_AMT        As Single           '现金支付金额
'        HAI_INBED_SBS_AMT   As Single           '补充住院补助支付金额
'        HAI_INBED_AMT       As Single           '补充住院支付金额
'        HAI_INBED_REPAY_AMT As Single           '补充住院再次支付金额
'        HAI_INBED_HIGH_AMT  As Single           '补充住院高额支付金额
'        OFFICAL_HIGH_AMT    As Single           '公务员高额补助支付金额
'        OFFICAL_INBED_AMT   As Single           '公务员住院补助支付金额
'        OFFICAL_ACCT_AMT    As Single           '公务员个帐补助支付金额

        .个人帐户 = IIf(cur帐户支付 = 0, gInBedBillInfo_乐山.INBED_ACCT_AMT, cur帐户支付)
        .补充基金 = gInBedBillInfo_乐山.HAI_INBED_SBS_AMT + gInBedBillInfo_乐山.HAI_INBED_AMT + _
                    gInBedBillInfo_乐山.HAI_INBED_REPAY_AMT + gInBedBillInfo_乐山.HAI_INBED_HIGH_AMT
        .医保基金 = gInBedBillInfo_乐山.INBED_FUND_AMT
        .公务员基金 = gInBedBillInfo_乐山.OFFICAL_HIGH_AMT + _
                      gInBedBillInfo_乐山.OFFICAL_INBED_AMT + gInBedBillInfo_乐山.OFFICAL_ACCT_AMT
        .医保总费用 = gInBedBillInfo_乐山.PART_PUB_AMT + gInBedBillInfo_乐山.PART_SELF_AMT + _
                      gInBedBillInfo_乐山.PUB_PAY_AMT + gInBedBillInfo_乐山.SELF_PAY_AMT
        If cur帐户支付 <> 0 Then
            .现金 = .总费用 - .医保基金 - .补充基金 - .个人帐户 - .公务员基金
        End If
    End With
End Sub

Public Function 上传处方_乐山(ByVal int性质 As Integer, ByVal int状态 As Integer, ByVal str单据号 As String) As Boolean
    Dim intType As Integer
    Dim int特殊项目 As Integer          '表明该工作站输入中心允许的特殊项目时,这类项目都做为特殊项目上传
    Dim int可做为特殊项目 As Integer    '表明该项目,中心是否允许做为特殊项目
    Dim lng病人ID As Long
    Dim blnInsure As Boolean, blnUpload As Boolean, blnTrans As Boolean
    Dim dbl单价 As Double
    Dim rsExse As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim gcn上传 As New ADODB.Connection
    On Error GoTo errHand
    
    Call DebugTool("提取处方明细")
    gstrSQL = " Select A.ID,A.病人ID,A.NO,A.序号,A.记录性质,A.记录状态,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 登记时间,A.收费类别," & _
              " A.开单人 医生,B.名称 开单部门,A.收费细目ID,D.名称 细目名称,C.项目编码 医保项目编码,C.医保大类,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传,A.摘要" & _
              " From 住院费用记录 A,部门表 B,收费细目 D," & _
              "     (Select A.*,B.名称 医保大类 From 保险支付项目 A,保险支付大类 B " & _
              "     Where A.险类=B.险类 And A.大类ID=B.ID And A.险类=" & TYPE_乐山 & ") C,保险帐户 G " & _
              " Where A.记录性质=[1] And A.记录状态=[2] And A.NO=[3]" & _
              " And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And A.收费细目ID=D.ID And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0" & _
              " And A.病人ID=G.病人ID And G.险类=[4]" & _
              " Order by A.NO,A.病人ID"
    Set rsExse = zlDatabase.OpenSQLRecord(gstrSQL, "读取费用明细", int性质, int状态, str单据号, TYPE_乐山)
    
    Set gcn上传 = GetNewConnection
    With rsExse
        Do While Not .EOF
            If lng病人ID <> !病人ID Then
                '提交数据
                If lng病人ID <> 0 And blnInsure Then
                    Call DebugTool("保存单据")
                    gbytReturn_乐山 = LS_SaveDetail(gtypBalance.顺序号)
                    If GetErrInfo_乐山 Then
                        gcn上传.RollbackTrans
                        上传处方_乐山 = True
                        Exit Function
                    End If
                    gcn上传.CommitTrans
                    blnTrans = False
                End If
            End If
            
            '判断当前病人是否本次以医保身份登记
            If lng病人ID <> !病人ID Then
                Call DebugTool("判断是否为医保病人")
                blnInsure = 是否医保病人(!病人ID)
            End If
            If blnInsure Then
                If lng病人ID <> !病人ID Then
                    Call DebugTool("获取该病人基本信息及住院信息")
                    lng病人ID = !病人ID
                    Call 获取病人基本信息(lng病人ID)
                    gbytReturn_乐山 = LS_GetInHospitalRegInfo(gtypBalance.顺序号)
                    gcn上传.BeginTrans
                    blnTrans = True
                    If GetErrInfo_乐山 Then
                        gcn上传.RollbackTrans
                        上传处方_乐山 = True
                        Exit Function
                    End If
                End If
                
                '上传明细
                intType = 1
                If !医保大类 = "诊疗" Then intType = 2
                If !医保大类 = "服务" Then intType = 3
                
'                Call DebugTool("判断可否做为特殊项目")
'                int可做为特殊项目 = 0
'                gstrSQL = "Select 附注 From 保险支付项目 Where 险类=" & TYPE_乐山 & " And 收费细目ID=" & !收费细目ID
'                Call OpenRecordset(rsTemp, "判断医保项目是否允许做为特殊项目")
'                If Not rsTemp.EOF Then
'                    If Not IsNull(rsTemp!附注) Then
'                        If UBound(Split(rsTemp!附注, "|")) >= 2 Then
'                            int可做为特殊项目 = Val(Split(rsTemp!附注, "|")(2))
'                        End If
'                    End If
'                End If
                
                '判断本笔明细是否做为特殊项目（因为录入费用明细时，就已经检查了每个项目是否允许做为特殊项目，此处不需要再次判断，只要摘要为1，说明做为特殊项目上传）
                int特殊项目 = IIf(Nvl(rsExse!摘要) = "1", 1, 0)
                
                '曾明春(2005-10-18) 判断是否设置医保编码，如果未设置，则进行提示
                If IsNull(rsExse!医保项目编码) Then
                   MsgBox "项目:" & rsExse!细目名称 & "[单价：" & rsExse!金额 / rsExse!数量 & "元]未设置医保编码!"
                   上传处方_乐山 = True
                   Exit Function
                End If
                
                Select Case intType
                Case 1
                    Call DebugTool("取药品信息")
                    gstrSQL = "select A.规格,A.产地,B.名称 剂型  " & _
                             " from 药品目录 A,药品剂型 B,药品信息 C " & _
                             " Where A.药名ID=C.药名ID And C.剂型=B.编码 And A.药品ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取药品信息", CLng(!收费细目ID))
                    
                    With gDrugInfo_乐山
                        .INBED_REG_ID = gtypBalance.顺序号.Data
                        .RECEIPT_DT = Format(rsExse!登记时间, "yyyy-MM-dd")
                        .DRUG_CATALOG_ID = rsExse!医保项目编码
                        .DRUG_INFO = rsExse!细目名称
                        dbl单价 = rsExse!金额 / rsExse!数量
                        '曾明春:2005-07-06 如果单价精度超过2位小数，则数量传1，单价传实收金额。
                        If Round(dbl单价 * 100) <> dbl单价 * 100 Then
                            '曾明春:2005-06-02修改，金额取绝对值,避免单价出现负数的情况
                            .UNIT_PRC = Format(Abs(rsExse!金额), "#####0.00;-#####0.00;0;")
                            If rsExse!金额 <= 0 Then
                              .SRVC_CNT = -1
                            Else
                              .SRVC_CNT = 1
                            End If
                        Else
                            .UNIT_PRC = Format(rsExse!金额 / rsExse!数量, "#####0.00;-#####0.00;0;")
                            .SRVC_CNT = rsExse!数量
                        End If
                        .COST_PRC = 0
                        .DRUG_TYP = Nvl(rsTemp!剂型)
                        .DRUG_SPEC = Nvl(rsTemp!规格)
                        .PRODUCE_FACTORY = Nvl(rsTemp!产地)
                        .DRUG_SPC_FLAG = int特殊项目
                    End With
                Case 2
                    Call DebugTool("取诊疗信息")
                    With gDiagInfo_乐山
                        .INBED_REG_ID = gtypBalance.顺序号.Data
                        .RECEIPT_DT = Format(rsExse!登记时间, "yyyy-MM-dd")
                        .DIAG_CATALOG_ID = rsExse!医保项目编码
                        .DIAG_ITEM_NAME = rsExse!细目名称
                        dbl单价 = rsExse!金额 / rsExse!数量
                        '曾明春:2006-11-06 如果单价精度超过2位小数，则数量传1，单价传实收金额。
                        If Round(dbl单价 * 100) <> dbl单价 * 100 Then
                           '曾明春:2006-11-06修改，金额取绝对值,避免单价出现负数的情况
                            .UNIT_PRC = Format(Abs(rsExse!金额), "#####0.00;-#####0.00;0;")
                            If rsExse!金额 <= 0 Then
                              .SRVC_CNT = -1
                            Else
                              .SRVC_CNT = 1
                            End If
                        Else
                            .UNIT_PRC = Format(rsExse!金额 / rsExse!数量, "#####0.00;-#####0.00;0;")
                            .SRVC_CNT = rsExse!数量
                        End If
                    End With
                Case 3
                    Call DebugTool("取服务设施信息")
                    With gServiceItemInfo_乐山
                        .INBED_REG_ID = gtypBalance.顺序号.Data
                        .RECEIPT_DT = Format(rsExse!登记时间, "yyyy-MM-dd")
                        .SRVC_ITEM_ID = rsExse!医保项目编码
                        .SRVC_NAME = rsExse!细目名称
                        dbl单价 = rsExse!金额 / rsExse!数量
                        '曾明春:2006-11-06 如果单价精度超过2位小数，则数量传1，单价传实收金额。
                        If Round(dbl单价 * 100) <> dbl单价 * 100 Then
                           '曾明春:2006-11-06修改，金额取绝对值,避免单价出现负数的情况
                            .UNIT_PRC = Format(Abs(rsExse!金额), "#####0.00;-#####0.00;0;")
                            If rsExse!金额 <= 0 Then
                              .SRVC_CNT = -1
                            Else
                              .SRVC_CNT = 1
                            End If
                        Else
                            .UNIT_PRC = Format(rsExse!金额 / rsExse!数量, "#####0.00;-#####0.00;0;")
                            .SRVC_CNT = rsExse!数量
                        End If
                    End With
                End Select
                
                Call DebugTool("上传明细")
                If Not UploadDetail(intType) Then
                    gcn上传.RollbackTrans
                    上传处方_乐山 = True
                    Exit Function
                End If
                
                '打上标记
                Call DebugTool("打上传标记")
                '仅打上传标志，因为明细必须正确上传后，才能保证结算的正确性
                'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
                gcn上传.Execute gstrSQL, , adCmdStoredProc
                Call DebugTool("上传标记SQL：" & gstrSQL)
                
                blnUpload = True
            End If
            .MoveNext
        Loop
        
       
        Call DebugTool("保存上传明细")
        If blnUpload And blnInsure Then
            gbytReturn_乐山 = LS_SaveDetail(gtypBalance.顺序号)
            If GetErrInfo_乐山 Then
                gcn上传.RollbackTrans
                上传处方_乐山 = True
                Exit Function
            End If
            Call DebugTool("提交数据")
            gcn上传.CommitTrans
            blnTrans = False
        End If
    End With
    
    上传处方_乐山 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    上传处方_乐山 = True
    If blnTrans Then gcn上传.RollbackTrans
End Function

Private Function UploadDetail(Optional ByVal intType As Integer = 1, Optional ByVal bln住院 As Boolean = True) As Boolean
    '上传费用明细
    'intType:1-药品;2-诊疗;3-服务
    If bln住院 Then
        Select Case intType
        Case 1
            gbytReturn_乐山 = LS_AddDrug(gDrugInfo_乐山)
        Case 2
            gbytReturn_乐山 = LS_AddDiag(gDiagInfo_乐山)
        Case 3
            gbytReturn_乐山 = LS_AddService(gServiceItemInfo_乐山)
        End Select
    Else
        Select Case intType
        Case 1
            gbytReturn_乐山 = LS_ExamAddDrug(gDrugInfo_乐山)
        Case 2
            gbytReturn_乐山 = LS_ExamAddDiag(gDiagInfo_乐山)
        Case 3
            gbytReturn_乐山 = LS_ExamAddServiceItem(gServiceItemInfo_乐山)
        End Select
    End If
    If GetErrInfo_乐山 Then Exit Function
    UploadDetail = True
End Function

Private Function TrimTsChar(ByVal strData As Variant) As String
    TrimTsChar = Replace(Replace(strData, " ", ""), Chr(0), "")
End Function
