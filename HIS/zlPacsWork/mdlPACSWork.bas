Attribute VB_Name = "mdlPACSWork"
Option Explicit
Public gobjRegist As Object
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public glngModul As Long
Public glngSys As Long

Public gstrDBUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrUnitName As String '用户单位名称
Public gfrmMain As Object

Public gstrSQL As String
Public gblnOK As Boolean
Public glngTXTProc As Long
Public gbln加班加价 As Boolean
Public grsDuty As ADODB.Recordset '存放医生职务
Public grsSysPars As ADODB.Recordset

'医保变量
Public gclsInsure As New clsInsure

'CIS系统参数
Public gbln药品按规格下医嘱 As Boolean
Public gint过敏登记有效天数 As Integer
Public gbln长期医嘱次日生效 As Boolean
Public gbln药疗划价单 As Boolean
Public gbln其他划价单 As Boolean
Public gbln执行后审核 As Boolean

'HIS系统参数
Public gbln中医 As Boolean '是否使用中医病案
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"
Public gbytCardLen As String '就诊卡号长度
Public gblnCardHide As Boolean '就诊卡号密文显示
Public gstrCardMask As String  '就诊卡允许的字母前缀:AA|BB|CC...
Public gint挂号天数 As Integer '挂号单有效天数
Public gbln收费类别 As Boolean '是否首先输入类别
Public gbln商品名 As Boolean '西成药是否按商品名显示
Public gbln住院自动发料 As Boolean '住院记帐完成后是否自动发料
Public gbln门诊自动发料 As Boolean '门诊记帐完成后是否自动发料
Public gbln从项汇总折扣 As Boolean '从属项目汇总计算折扣
Public gbln报警包含划价费用 As Boolean '记帐报警包含划价费用

'医技工作站系统费用参数
Public gbytBillOpt As Byte '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
Public gblnStock As Boolean '指定药房时是否限定输入药品的库存
Public gstr医保费用类型 As String '医保病人允许的费用类型
Public gstr公费费用类型 As String '公费病人允许的费用类型

Public gintReportFormat As Integer     '记录报告打印格式
'-----------------------------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤消出院 = 17            '允许撤消病人出院
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
    support出院无实际交易 = 29 '出院接口中是否要与接口商进行交易
End Enum
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        " Select A.ID,C.部门ID,A.编号,A.简码,A.姓名,B.用户名" & _
        " From 人员表 A,上机人员表 B,部门人员 C" & _
        " Where A.ID = B.人员ID And A.ID = C.人员ID And C.缺省 = 1 And Upper(B.用户名) = USER"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, "mdlCureBase", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMaxBedLen(Optional lng部门ID As Long, Optional bln科室 As Boolean) As Integer
'功能：获取指定部门的床位号的最大长度
'参数：lng部门ID=病区ID或科室ID,为0表示所有病区或科室
'      bln占用=是否只管被占用的床
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not bln科室 Or lng部门ID = 0 Then
        strSQL = "Select Max(Length(床号)) as 长度 From 床位状况记录 Where 状态='占用' And 病区ID" & IIf(lng部门ID = 0, " is Not NULL", "=" & lng部门ID)
    Else
        strSQL = "Select Max(Length(床号)) as 长度 From 床位状况记录 Where 状态='占用' And 科室ID" & IIf(lng部门ID = 0, " is Not NULL", "=" & lng部门ID)
    End If
    
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, "mdlCISWork", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    If Not rsTmp.EOF Then GetMaxBedLen = IIf(IsNull(rsTmp!长度), 0, rsTmp!长度)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病区ID(lng科室ID As Long) As Long
'功能：从科室ID获取对应的病区ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 病区ID From 床位状况记录 Where 科室ID=[1] Group by 病区ID"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng科室ID)
    If Not rsTmp.EOF Then Get病区ID = rsTmp!病区ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser病区IDs() As String
'功能：获取操作员所属的病区(直接属于病区或所在科室所属的病区),可能有多个
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = _
        "Select Distinct 病区ID From (" & _
        " Select A.部门ID as 病区ID" & _
        " From 部门性质说明 A,部门人员 B" & _
        " Where A.部门ID=B.部门ID And B.人员ID=[1]" & _
        " And A.服务对象 in(1,2,3) And A.工作性质='护理'" & _
        " Union" & _
        " Select A.病区ID From 床位状况记录 A,部门人员 B" & _
        " Where A.科室ID=B.部门ID And B.人员ID=[1])"
    
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        GetUser病区IDs = GetUser病区IDs & "," & rsTmp!病区ID
        rsTmp.MoveNext
    Next
    GetUser病区IDs = Mid(GetUser病区IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser科室IDs(Optional ByVal bln病区 As Boolean) As String
'功能：获取操作员所属的科室(本身所在科室+所属病区包含的科室),可能有多个
'参数：是否取所属病区下的科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select 部门ID From 部门人员 Where 人员ID=[1]"
    If bln病区 Then
        strSQL = strSQL & " Union" & _
            " Select Distinct B.科室ID From 部门人员 A,床位状况记录 B" & _
            " Where A.部门ID=B.病区ID And A.人员ID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        GetUser科室IDs = GetUser科室IDs & "," & rsTmp!部门ID
        rsTmp.MoveNext
    Next
    GetUser科室IDs = Mid(GetUser科室IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitStockCheck(ByVal int范围 As Integer) As Collection
'功能：读取不同库房出库检查方式于集合中
'参数：int范围=1-门诊,2-住院
    Dim colStock As Collection
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    '不同药房药品出库检查方式
    Set colStock = New Collection
    colStock.Add 0, "_0" '加一条,防万一出错
    
    strSQL = _
        " Select Distinct A.ID,C.检查方式" & _
        " From 部门表 A,部门性质说明 B,药品出库检查 C" & _
        " Where B.部门ID=A.ID And B.服务对象 IN([1],3)" & _
        " And B.工作性质 in('中药房','西药房','成药房')" & _
        " And C.库房ID(+)=A.ID"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int范围)
    For i = 1 To rsTmp.RecordCount
        colStock.Add Nvl(rsTmp!检查方式, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    Set InitStockCheck = colStock
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    
    On Error GoTo errH
        
    'HIS系统参数
    '---------------------------------------------------------
    strSQL = "Select 参数号,参数名,参数值 from 系统参数表"
    Call OpenRecord(rsTmp, strSQL, "mdlCISWork")
    
    '费用金额小数点位数
    gbytDec = 2: gstrDec = "0.00"
    rsTmp.Filter = "参数号=9"
    If Not rsTmp.EOF Then
        gbytDec = Val(Nvl(rsTmp!参数值, 2))
        gstrDec = "0." & String(gbytDec, "0")
    End If
    
    '就诊卡号密文显示
    rsTmp.Filter = "参数号=12"
    If Not rsTmp.EOF Then gblnCardHide = Nvl(rsTmp!参数值, 0) <> 0
    
    '指定药房时限制库存
    rsTmp.Filter = "参数号=18"
    If Not rsTmp.EOF Then gblnStock = Nvl(rsTmp!参数值, 0) <> 0
    
    '就诊卡号码的长度
    gbytCardLen = 7
    rsTmp.Filter = "参数号=20"
    If Not rsTmp.EOF Then
        gbytCardLen = Val(Split(Nvl(rsTmp!参数值, "7|7|7|7|7"), "|")(4))
    End If
    
    '挂号有效天数
    rsTmp.Filter = "参数号=21"
    If Not rsTmp.EOF Then gint挂号天数 = Nvl(rsTmp!参数值, 0)
    
    '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
    rsTmp.Filter = "参数号=23"
    If Not rsTmp.EOF Then gbytBillOpt = Nvl(rsTmp!参数值, 0)
    
    '就诊卡识别前缀
    rsTmp.Filter = "参数号=27"
    If Not rsTmp.EOF Then gstrCardMask = UCase(Nvl(rsTmp!参数值))
    
    '是否使用中医
    rsTmp.Filter = "参数号=31"
    If Not rsTmp.EOF Then gbln中医 = Nvl(rsTmp!参数值, 0) <> 0
    
    '医保费用类型
    rsTmp.Filter = "参数号=41"
    If Not rsTmp.EOF Then
        gstr医保费用类型 = "'" & Replace(Nvl(rsTmp!参数值), "|", "','") & "'"
    End If

    '公费费用类型
    rsTmp.Filter = "参数号=42"
    If Not rsTmp.EOF Then
        gstr公费费用类型 = "'" & Replace(Nvl(rsTmp!参数值), "|", "','") & "'"
    End If
    
    '住院自动发料
    rsTmp.Filter = "参数号=63"
    If Not rsTmp.EOF Then
        gbln住院自动发料 = Nvl(rsTmp!参数值, 0) <> 0
    End If
    
    '药品按规格下医嘱
    rsTmp.Filter = "参数号=69"
    If Not rsTmp.EOF Then gbln药品按规格下医嘱 = Val(Nvl(rsTmp!参数值, 0)) = 1
    
    '皮试结果有效时间
    rsTmp.Filter = "参数号=70"
    If Not rsTmp.EOF Then gint过敏登记有效天数 = Val(Nvl(rsTmp!参数值, 0))
    
    '长期医嘱次日生效
    rsTmp.Filter = "参数号=71"
    If Not rsTmp.EOF Then gbln长期医嘱次日生效 = Val(Nvl(rsTmp!参数值, 0)) = 1
    
    '是否要求首先输入类别
    rsTmp.Filter = "参数号=72"
    If Not rsTmp.EOF Then gbln收费类别 = Nvl(rsTmp!参数值, 1) <> 0
    
    '西成药是否按商品名显示
    rsTmp.Filter = "参数号=74"
    If Not rsTmp.EOF Then gbln商品名 = Nvl(rsTmp!参数值, 0) <> 0
    
    '药疗生成划价单
    rsTmp.Filter = "参数号=79"
    If Not rsTmp.EOF Then gbln药疗划价单 = Nvl(rsTmp!参数值, 0) <> 0
    
    '其他生成划价单
    rsTmp.Filter = "参数号=80"
    If Not rsTmp.EOF Then gbln其他划价单 = Nvl(rsTmp!参数值, 0) <> 0
    
    '执行后自动审核
    rsTmp.Filter = "参数号=81"
    If Not rsTmp.EOF Then gbln执行后审核 = Nvl(rsTmp!参数值, 0) <> 0
            
    '门诊自动发料
    rsTmp.Filter = "参数号=92"
    If Not rsTmp.EOF Then
        gbln门诊自动发料 = Nvl(rsTmp!参数值, 0) <> 0
    End If
    
    '从属项目汇总计算折扣
    rsTmp.Filter = "参数号=93"
    If Not rsTmp.EOF Then gbln从项汇总折扣 = Nvl(rsTmp!参数值, 0) <> 0
    
    '记帐报警包含划价费用
    rsTmp.Filter = "参数号=98"
    If Not rsTmp.EOF Then gbln报警包含划价费用 = Nvl(rsTmp!参数值, 0) <> 0
    
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    InitSysPar = False
End Function

Public Function GetPatiYear(lng病人id As Long) As Integer
'功能：获取病人的准确年龄
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intYear As Integer
    
    On Error GoTo errH
    
    strSQL = "Select Sysdate as 当前,出生日期,年龄 From 病人信息 Where 病人ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng病人id)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!出生日期) Then
            intYear = Year(rsTmp!当前) - Year(rsTmp!出生日期)
            If Format(rsTmp!当前, "MMdd") < Format(rsTmp!出生日期, "MMdd") Then
                intYear = intYear - 1
            End If
            If intYear < 0 Then intYear = 0
        Else
            intYear = Val(Nvl(rsTmp!年龄))
        End If
    End If
    GetPatiYear = intYear
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get部门名称(lng部门ID As Long) As String
'功能：返回部门名称
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 名称 From 部门表 Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng部门ID)
    If Not rsTmp.EOF Then Get部门名称 = rsTmp!名称
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get项目名称(lng项目ID As Long) As String
'功能：返回诊疗项目名称
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 名称 From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
    If Not rsTmp.EOF Then Get项目名称 = rsTmp!名称
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get缺省用法ID(ByVal int类型 As Integer, ByVal int来源 As Integer) As Long
'功能：返回缺省的给药途径或中药煎法
'参数：int类型=2-给药途径,3-中药煎法,4-中药用法,6-采集方法(检验)
'      int来源=1-门诊,2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID From 诊疗项目目录" & _
        " Where 类别='E' And 操作类型=[1]" & _
        " And (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
        " And 服务对象 IN([2],3) And Rownum<100" & _
        " Order by 编码"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", CStr(int类型), int来源)
    If Not rsTmp.EOF Then
        Get缺省用法ID = rsTmp!ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check上班安排(ByVal bln药房 As Boolean) As Boolean
'功能：检查医院的科室是否使用了上班安排
'参数：bln药房=是检查药房上班还是其它科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If bln药房 Then
        strSQL = "Select Count(B.部门ID) as NUM From 部门性质说明 A,部门安排 B" & _
            " Where A.部门ID=B.部门ID And A.工作性质 IN('西药房','成药房','中药房')"
    Else
        strSQL = "Select Count(B.部门ID) as NUM From 部门性质说明 A,部门安排 B" & _
            " Where A.部门ID=B.部门ID And A.工作性质 Not IN('西药房','成药房','中药房')"
    End If
    Call OpenRecord(rsTmp, strSQL, "mdlCISWork")
    If Not rsTmp.EOF Then
        Check上班安排 = Nvl(rsTmp!Num, 0) > 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get收费执行科室ID(ByVal str类别 As String, ByVal lng项目ID As Long, _
    ByVal int执行科室 As Integer, ByVal lng科室ID As Long, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal lng发料部门 As Long) As Long
'功能：获取非药收费项目的执行科室
'参数：int范围=1.门诊,2-住院
'      lng科室ID=病人科室ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim str药房 As String, lng药房 As Long
    Dim bytDay As Byte, strIDs As String
    
    On Error GoTo errH
    
    If str类别 = "4" Then
        '以及SQL在卫材不支持存储库房设置之前用
'        strSQL = "Select B.服务对象,A.编码,A.ID From 部门表 A,部门性质说明 B" & _
'            " Where A.ID=B.部门ID And B.工作性质='发料部门' And B.服务对象 IN([1],3)" & _
'            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
'            " Order by B.服务对象,A.编码"
'        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int范围)
'        If Not rsTmp.EOF Then
'            If lng发料部门 <> 0 Then rsTmp.Filter = "ID=" & lng发料部门
'            If rsTmp.EOF Then rsTmp.Filter = 0
'            Get收费执行科室ID = rsTmp!ID
'        End If
        
        strSQL = _
            " Select Distinct" & _
            "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            " And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
            " And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int范围, lng科室ID, lng项目ID)
        If Not rsTmp.EOF Then
            rsTmp.Filter = "开单科室ID=" & lng科室ID
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0"
            For i = 1 To rsTmp.RecordCount
                If i = 1 Or rsTmp!执行科室ID = lng发料部门 Then
                    Get收费执行科室ID = rsTmp!执行科室ID
                End If
                rsTmp.MoveNext
            Next
        End If
    ElseIf InStr(",5,6,7,", str类别) > 0 Then
        If str类别 = "5" Then
            str药房 = "西药房"
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIf(int范围 = 2, "住院", "门诊") & "缺省西药房", 0))
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIf(int范围 = 2, "住院", "门诊") & "缺省成药房", 0))
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIf(int范围 = 2, "住院", "门诊") & "缺省中药房", 0))
        End If
        
        '药品从系统指定的储备药房中找
        If Not Check上班安排(True) Then
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And A.收费细目ID=[4]" & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And A.收费细目ID=[4]" & _
                " Order by B.服务对象,C.编码"
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str药房, int范围, lng科室ID, lng项目ID, bytDay)
        If Not rsTmp.EOF Then
            strIDs = ""
            rsTmp.Filter = "开单科室ID=" & lng科室ID
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0"
            For i = 1 To rsTmp.RecordCount
                strIDs = strIDs & "," & rsTmp!执行科室ID '收集用于动态分配
                If i = 1 Or rsTmp!执行科室ID = lng药房 Then
                    Get收费执行科室ID = rsTmp!执行科室ID
                    If rsTmp!执行科室ID = lng药房 Then
                        strIDs = "": Exit For
                    End If
                End If
                rsTmp.MoveNext
            Next
            strIDs = Mid(strIDs, 2)
            If UBound(Split(strIDs, ",")) <= 0 Then strIDs = ""
        End If
    Else
        Select Case int执行科室
            Case 0 '0-无明确科室
                Get收费执行科室ID = UserInfo.部门ID
            Case 1 '1-病人所在科室
                Get收费执行科室ID = lng科室ID
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    Get收费执行科室ID = lng科室ID
                Else
                    Get收费执行科室ID = Get病区ID(lng科室ID)
                End If
            Case 3 '3-开单人所在科室
                Get收费执行科室ID = UserInfo.部门ID
            Case 4 '4-指定科室
                strSQL = "Select Nvl(开单科室ID,0) as 开单科室ID,执行科室ID" & _
                    " From 收费执行科室 Where 收费细目ID=[1]" & _
                    " And (病人来源 is NULL Or 病人来源=[2])" & _
                    " And (开单科室ID is NULL Or 开单科室ID=[3])" & _
                    " Order by Decode(病人来源,Null,2,1)" '默认科室优先
                Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID, int范围, lng科室ID)
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "开单科室ID=" & lng科室ID
                    If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0"
                    If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
                End If
        End Select
        If Get收费执行科室ID = 0 Then Get收费执行科室ID = UserInfo.部门ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get诊疗执行科室ID(ByVal str类别 As String, ByVal lng项目ID As Long, _
    ByVal lng药品ID As Long, ByVal int执行科室 As Integer, ByVal lng科室ID As Long, _
    ByVal int期效 As Integer, Optional ByVal int范围 As Integer = 2) As Long
'功能：根据诊疗项目执行科室信息返回缺省的执行科室ID
'参数：lng药品ID=药品ID,确定到规格时要用
'      int执行科室=项目执行科室标志
'      lng科室ID=病人科室ID
'      lng西药房,lng成药房,lng中药房=药品缺省药房,药品类时需要
'      int期效=0-长嘱,1-临嘱,临嘱可能要判断上班时间
'      int范围=1-门诊,2-住院(缺省)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str药房 As String, lng药房 As Long
    Dim bytDay As Byte, bln上班安排 As Boolean
    Dim bln规格 As Boolean
    
    On Error GoTo errH
    
    If InStr(",5,6,7,", str类别) > 0 Then
        bln规格 = ((int期效 = 1 Or gbln药品按规格下医嘱) And lng药品ID <> 0) Or lng药品ID <> 0
        
        If str类别 = "5" Then
            str药房 = "西药房"
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIf(int范围 = 2, "住院", "门诊") & "缺省西药房", 0))
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIf(int范围 = 2, "住院", "门诊") & "缺省成药房", 0))
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIf(int范围 = 2, "住院", "门诊") & "缺省中药房", 0))
        End If
        
        '药品从系统指定的储备药房中找
        If int范围 = 1 Then bln上班安排 = Check上班安排(True) '住院医嘱不管药房上班安排
        If Not bln上班安排 Then
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From " & IIf(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                 IIf(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From " & IIf(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[6]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                IIf(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & _
                " Order by B.服务对象,C.编码"
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str药房, int范围, lng科室ID, lng药品ID, lng项目ID, bytDay)
        If Not rsTmp.EOF Then
            rsTmp.Filter = "开单科室ID=" & lng科室ID
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0"
            For i = 1 To rsTmp.RecordCount
                If i = 1 Or rsTmp!执行科室ID = lng药房 Then
                    Get诊疗执行科室ID = rsTmp!执行科室ID
                    If rsTmp!执行科室ID = lng药房 Then Exit For
                End If
                rsTmp.MoveNext
            Next
        End If
    Else
        Select Case int执行科室
            Case 0 '0-无执行的叮嘱
                Exit Function
            Case 1 '1-病人所在科室
                Get诊疗执行科室ID = lng科室ID
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    Get诊疗执行科室ID = lng科室ID
                Else
                    Get诊疗执行科室ID = Get病区ID(lng科室ID)
                End If
            Case 3 '3-开单人所在科室
                Get诊疗执行科室ID = UserInfo.部门ID
            Case 4 '4-指定科室
                If int期效 = 1 Then bln上班安排 = Check上班安排(False)
                If Not bln上班安排 Then
                    strSQL = "Select Nvl(开单科室ID,0) as 开单科室ID,执行科室ID" & _
                        " From 诊疗执行科室" & _
                        " Where 诊疗项目ID=[1]" & _
                        " And (病人来源 is NULL Or 病人来源=[2])" & _
                        " And (开单科室ID is NULL Or 开单科室ID=[3])" & _
                        " Order by Decode(病人来源,Null,2,1)" '默认科室优先
                Else
                    bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
                    strSQL = _
                        " Select Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                        " From 诊疗执行科室 A,部门安排 B" & _
                        " Where A.执行科室ID+0=B.部门ID And B.星期=[4]" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.开始时间,'HH24:MI:SS') and To_Char(B.终止时间,'HH24:MI:SS') " & _
                        " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                        " And A.诊疗项目ID=[1]" & _
                        " Order by Decode(A.病人来源,Null,2,1)"
                End If
                Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID, int范围, lng科室ID, bytDay)
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "开单科室ID=" & lng科室ID
                    If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0"
                    If Not rsTmp.EOF Then Get诊疗执行科室ID = rsTmp!执行科室ID
                End If
            Case 5 '5-院外执行
                Exit Function
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get可用药房IDs(ByVal str类别 As String, ByVal lng项目ID As Long, _
    ByVal lng药品ID As Long, ByVal lng科室ID As Long, Optional ByVal int范围 As Integer = 2) As String
'功能：获取药品的有效诊疗执行科室ID串,用于判断缺省执行科室
'参数：lng科室ID=病人科室ID
'      int范围=1-门诊,2-住院(缺省)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str药房 As String
    Dim bytDay As Byte, bln上班安排 As Boolean
    Dim str药房IDs As String
    
    '系统可以指定药品执行科室,这里提取所有可选的供再选择
    If str类别 = "5" Then
        str药房 = "西药房"
    ElseIf str类别 = "6" Then
        str药房 = "成药房"
    ElseIf str类别 = "7" Then
        str药房 = "中药房"
    End If
        
    '药品从系统指定的储备药房中找
    If int范围 = 1 Then bln上班安排 = Check上班安排(True) '住院医嘱不管药房上班安排
    If Not bln上班安排 Then
        strSQL = _
            " Select Distinct C.ID" & _
            " From " & IIf(lng药品ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
            " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
            IIf(lng药品ID <> 0, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]")
    Else
        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
        strSQL = _
            " Select Distinct C.ID" & _
            " From " & IIf(lng药品ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C,部门安排 D" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
            " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And D.部门ID=C.ID And D.星期=[6]" & _
            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
            " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
            IIf(lng药品ID <> 0, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]")
    End If
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str药房, int范围, lng科室ID, lng药品ID, lng项目ID, bytDay)
    Do While Not rsTmp.EOF
        str药房IDs = str药房IDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    Get可用药房IDs = Mid(str药房IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get诊疗执行科室(objCbo As Object, ByVal str类别 As String, ByVal lng项目ID As Long, ByVal lng药品ID As Long, _
    ByVal int执行科室 As Integer, ByVal lng科室ID As Long, ByVal lng当前执行ID As Long, ByVal int期效 As Integer, Optional ByVal int范围 As Integer = 2) As Boolean
'功能：根据诊疗项目执行科室信息返回可用的执行科室在指定下拉框中
'参数：int执行科室=项目执行科室标志
'      lng科室ID=病人科室ID
'      lng当前执行ID=医嘱当前的执行科室ID
'      int期效=0-长嘱,1-临嘱,临嘱可能要判断上班时间
'      int范围=1-门诊,2-住院(缺省)
'说明：对非药医嘱,当前的执行科室可能是强行选择出来的,需要显示在选择框中;另选择框中增加一个其它供选择
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str药房 As String
    Dim bytDay As Byte, bln上班安排 As Boolean
    Dim bln规格 As Boolean, i As Long
    
    If InStr(",5,6,7,", str类别) > 0 Then
        bln规格 = ((int期效 = 1 Or gbln药品按规格下医嘱) And lng药品ID <> 0) Or lng药品ID <> 0
        
        '系统可以指定药品执行科室,这里提取所有可选的供再选择
        If str类别 = "5" Then
            str药房 = "西药房"
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
        End If
            
        '药品从系统指定的储备药房中找
        If int范围 = 1 Then bln上班安排 = Check上班安排(True) '住院医嘱不管药房上班安排
        If Not bln上班安排 Then
            strSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象" & _
                " From " & IIf(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                IIf(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象" & _
                " From " & IIf(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[8]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                IIf(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & _
                " Order by B.服务对象,C.编码"
        End If
    Else
        Select Case int执行科室
            Case 0, 5 '0-无执行的叮嘱,5-院外执行
                Get诊疗执行科室 = True: Exit Function
            Case 1 '1-病人所在科室
                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([3],[6]) Order by 编码"
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([3],[6]) Order by 编码"
                Else
                    strSQL = _
                        " Select A.ID,A.编码,A.简码,A.名称" & _
                        " From 部门表 A,床位状况记录 B" & _
                        " Where Rownum<2 And A.ID=B.病区ID And B.科室ID=[3]" & _
                        " Union " & _
                        " Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                        " Order by 编码"
                End If
            Case 3 '3-开单人所在科室
                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([7],[6]) Order by 编码"
            Case 4 '4-指定科室
                If int期效 = 1 Then bln上班安排 = Check上班安排(False)
                If Not bln上班安排 Then
                    strSQL = _
                        " Select Distinct A.ID,A.编码,A.简码,A.名称" & _
                        " From 部门表 A,诊疗执行科室 B" & _
                        " Where A.ID=B.执行科室ID And B.诊疗项目ID=[5]" & _
                        " And (B.病人来源 is NULL Or B.病人来源=[2])" & _
                        " And (B.开单科室ID is NULL Or B.开单科室ID=[3])" & _
                        " Union Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                        " Order by 编码"
                Else
                    bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.简码,C.名称" & _
                        " From 诊疗执行科室 A,部门安排 B,部门表 C" & _
                        " Where A.执行科室ID+0=B.部门ID And B.部门ID=C.ID And B.星期=[8]" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.开始时间,'HH24:MI:SS') and To_Char(B.终止时间,'HH24:MI:SS') " & _
                        " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                        " And A.诊疗项目ID=[5]" & _
                        " Union Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                        " Order by 编码"
                End If
        End Select
    End If
        
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str药房, int范围, lng科室ID, lng药品ID, lng项目ID, lng当前执行ID, UserInfo.部门ID, bytDay)
    objCbo.Clear
    For i = 1 To rsTmp.RecordCount
        '使用API快速加入,不然可能有点慢
        AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, rsTmp!编码 & "-" & rsTmp!名称
        SetComboData objCbo.Hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
        If lng当前执行ID = rsTmp!ID Then
            Call zlControl.CboSetIndex(objCbo.Hwnd, i - 1)
        End If
        rsTmp.MoveNext
    Next
    
    '仅非药医嘱可以选择
    If InStr(",5,6,7,", str类别) = 0 Then
        AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, "[其它...]"
        SetComboData objCbo.Hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get频率信息_编码(ByVal str编码 As String, str频率 As String, _
    int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String) As Boolean
'功能：返回频率的相关信息
'参数：str编码=频率编码
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    str频率 = ""
    int频率次数 = 0
    int频率间隔 = 0
    str间隔单位 = ""
    
    strSQL = "Select * From 诊疗频率项目 Where 编码=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str编码)
    If Not rsTmp.EOF Then
        str频率 = Nvl(rsTmp!名称)
        int频率次数 = Nvl(rsTmp!频率次数, 0)
        int频率间隔 = Nvl(rsTmp!频率间隔, 0)
        str间隔单位 = Nvl(rsTmp!间隔单位)
    End If
    Get频率信息_编码 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get频率信息_名称(ByVal str频率 As String, int频率次数 As Integer, _
    int频率间隔 As Integer, str间隔单位 As String, str范围 As String) As Boolean
'功能：返回频率的相关信息
'参数：str频率=频率名称
'      str范围=1-西医,2-中医,-1-一次性,-2-持续性
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    int频率次数 = 0
    int频率间隔 = 0
    str间隔单位 = ""
    
    strSQL = "Select * From 诊疗频率项目 Where 名称=[1] And Instr([2],','||适用范围||',')>0"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str频率, "," & str范围 & ",")
    If Not rsTmp.EOF Then
        int频率次数 = Nvl(rsTmp!频率次数, 0)
        int频率间隔 = Nvl(rsTmp!频率间隔, 0)
        str间隔单位 = Nvl(rsTmp!间隔单位)
    End If
    Get频率信息_名称 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get缺省频率(ByVal int范围 As Integer, str频率 As String, _
    int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String) As Boolean
'功能：从所有适用频率项目中取一个作为缺省频率
'参数：str范围=1-西医,2-中医,-1-一次性,-2-持续性
'返回：缺省频率信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    str频率 = ""
    int频率次数 = 0
    int频率间隔 = 0
    str间隔单位 = ""
    
    strSQL = "Select * From 诊疗频率项目 Where 适用范围=[1] Order by 编码"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int范围)
    If Not rsTmp.EOF Then
        str频率 = Nvl(rsTmp!名称)
        int频率次数 = Nvl(rsTmp!频率次数, 0)
        int频率间隔 = Nvl(rsTmp!频率间隔, 0)
        str间隔单位 = Nvl(rsTmp!间隔单位)
    End If
    Get缺省频率 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get诊疗频率(int范围 As Integer, objCbo As Object, Optional str频率 As String) As Boolean
'功能：读取诊疗频率项目在指定下拉框中,并设置缺省项
'参数：int范围=1-西医,2-中医
'      str频率=缺省频率名称
'说明：非药品可选频率项目使用西药类的
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    strSQL = "Select 英文名称,名称 From 诊疗频率项目 Where 适用范围=[1] Order by 编码"

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int范围)
    objCbo.Clear
    For i = 1 To rsTmp.RecordCount
        objCbo.AddItem Nvl(rsTmp!英文名称) & "-" & rsTmp!名称
        If str频率 = rsTmp!名称 Then
            Call zlControl.CboSetIndex(objCbo.Hwnd, objCbo.NewIndex)
        End If
        rsTmp.MoveNext
    Next
    Get诊疗频率 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get缺省时间(int范围 As Integer, str频率 As String, Optional lng给药途径ID As Long) As String
'功能：获取指定执行频率缺省的执行时间方案
'参数：int范围=1-西医;2-中医;-1-一次性;-2-持续性
'      lng给药途径ID=是否可以按指定给药途径优先取,否则取不确定给药途径的方案
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Nvl(A.给药途径ID,0) as 用法,A.时间方案" & _
        " From 诊疗频率时间 A,诊疗频率项目 B" & _
        " Where A.执行频率=B.编码 And B.适用范围=[1]" & _
        " And (A.给药途径ID is NULL Or A.给药途径ID=[2]) And B.名称=[3]" & _
        " Order by 用法 Desc,A.方案序号"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int范围, lng给药途径ID, str频率)
    If Not rsTmp.EOF Then Get缺省时间 = Nvl(rsTmp!时间方案)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get时间方案(objCbo As Object, int范围 As Integer, str频率 As String, Optional lng给药途径ID As Long) As Boolean
'功能：读取指定频率可用的诊疗频率时间方案在指定下拉框中,并设置缺省项(或保持原有值)
'参数：int范围=1-西医;2-中医;-1-一次性;-2-持续性
'      str频率=诊疗频率项目名称
'      lng给药途径ID=是否只读取指定给药途径的时间方案,否则为非药品的执行时间方案
'      str执行时间=缺省执行时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    '按理不同方案(不管是否指定给药途径)的执行时间应该不相同,否则会重复出现
    strSQL = "Select A.方案序号,A.时间方案" & _
        " From 诊疗频率时间 A,诊疗频率项目 B" & _
        " Where A.执行频率=B.编码 And B.名称=[1]" & _
        " And (A.给药途径ID is NULL Or A.给药途径ID=[2])" & _
        " And B.适用范围=[3]" & _
        " Order by A.方案序号"

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str频率, lng给药途径ID, int范围)
    strSQL = objCbo.Text: objCbo.Clear 'Clear会导致Text清空
    For i = 1 To rsTmp.RecordCount
        objCbo.AddItem rsTmp!时间方案
        rsTmp.MoveNext
    Next
    objCbo.Text = strSQL: objCbo.Tag = ""
    Get时间方案 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get开嘱医生(ByVal lng病人科室ID As Long, ByVal bln护士站 As Boolean, str缺省医生 As String, lng医生ID As Long, _
    Optional objCbo As Object, Optional ByVal int范围 As Integer = 2) As Boolean
'功能：获取可用的开嘱医生在指定的下拉框中
'参数：lng病人科室ID=病人所在科室ID
'      bln护士站=是否由护士代医生下医嘱
'      objCbo=要加入医生清单的下拉框
'      str缺省医生=缺省定位的医生,如果不传objCbo,则先优先定位,再返回缺省医生和医生ID
'      int范围=1-门诊,2-住院(缺省)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    If bln护士站 Then
        '病人所在科室的医生
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIf(objCbo Is Nothing, ",B.部门ID", "") & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID=[1]" & _
            " Order by A.简码"
        '病人所在病区各科的医生
        strSQL = "Select Distinct 病区ID From 床位状况记录 Where 科室ID=[1]"
        strSQL = "Select Distinct 科室ID From 床位状况记录 Where 病区ID=(" & strSQL & ")"
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIf(objCbo Is Nothing, ",B.部门ID", "") & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID IN(" & strSQL & ")" & _
            " Order by A.简码"
        '全院住院科室的医生
        strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN([2],3)"
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIf(objCbo Is Nothing, ",B.部门ID", "") & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID IN(" & strSQL & ")" & _
            " Order by A.简码"
    Else '医生下医嘱时,限制为只能为医生本人
        strSQL = "Select ID,编号,姓名,简码 From 人员表 Where ID=" & UserInfo.ID
    End If

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng病人科室ID, int范围)
    If objCbo Is Nothing Then
        If Not rsTmp.EOF Then
            If Not bln护士站 Then
                lng医生ID = rsTmp!ID
                str缺省医生 = rsTmp!姓名
            ElseIf bln护士站 Then
                If str缺省医生 <> "" Then
                    '缺省医生(住院医师)优先
                    rsTmp.Filter = "姓名='" & str缺省医生 & "'"
                Else
                    '病人科室的医生优先
                    rsTmp.Filter = "部门ID=" & lng病人科室ID
                End If
                If rsTmp.EOF Then rsTmp.Filter = 0
                lng医生ID = rsTmp!ID
                str缺省医生 = rsTmp!姓名
            End If
        End If
    Else
        objCbo.Clear
        For i = 1 To rsTmp.RecordCount
            AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, Nvl(rsTmp!简码) & "-" & rsTmp!姓名
            SetComboData objCbo.Hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
            If rsTmp!姓名 = str缺省医生 Then
                Call zlControl.CboSetIndex(objCbo.Hwnd, i - 1)
            End If
            'objCbo.AddItem Nvl(rsTmp!简码) & "-" & rsTmp!姓名
            'objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
'            If rsTmp!姓名 = str缺省医生 Then
'                Call zlControl.CboSetIndex(objCbo.hwnd, objCbo.NewIndex)
'            End If
            rsTmp.MoveNext
        Next
    End If
    Get开嘱医生 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get开嘱科室ID(ByVal lng医生ID As Long, ByVal lng病人科室ID As Long, Optional ByVal int范围 As Integer = 2) As Long
'功能：由医生确定开嘱科室
'参数：int范围=1-门诊,2-住院(缺省)
'说明：在医生所属科室范围内,优先顺序如下：
'      1、病人科室
'      2、服务于门诊/住院病人的科室且为默认科室
'      3、服务于门诊/住院病人的科室
'      4、默认科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr科室ID(1 To 4) As Long
    
    '可能部门没有性质
    strSQL = "Select Distinct C.编码,A.部门ID,Nvl(A.缺省,0) as 缺省,Nvl(B.服务对象,0) as 服务对象" & _
        " From 部门人员 A,部门性质说明 B,部门表 C" & _
        " Where A.部门ID=C.ID And A.部门ID=B.部门ID(+) And A.人员ID=[1]" & _
        " Order by C.编码"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医生ID)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!部门ID = lng病人科室ID Then
            arr科室ID(1) = rsTmp!部门ID
        ElseIf InStr("," & int范围 & ",3,", rsTmp!服务对象) > 0 And rsTmp!缺省 = 1 Then
            arr科室ID(2) = rsTmp!部门ID
        ElseIf InStr("," & int范围 & ",3,", rsTmp!服务对象) > 0 Then
            If arr科室ID(3) = 0 Then arr科室ID(3) = rsTmp!部门ID
        ElseIf rsTmp!缺省 = 1 Then
            arr科室ID(4) = rsTmp!部门ID
        End If
        rsTmp.MoveNext
    Next
    For i = LBound(arr科室ID) To UBound(arr科室ID)
        If arr科室ID(i) <> 0 Then
            Get开嘱科室ID = arr科室ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check本科执行(ByVal lng执行科室ID As Long) As Boolean
'功能：确定指定的执行科室是否本科(医生科室)
'参数：lng执行科室ID=医嘱的执行科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr科室ID(1 To 4) As Long
    
    strSQL = "Select 部门ID From 部门人员 Where 人员ID=[1] And 部门ID=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID, lng执行科室ID)
    Check本科执行 = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStock(ByVal lng药品ID As Long, ByVal lng库房ID As Long, Optional ByVal int范围 As Integer = 2) As Double
'功能：获取指定库房指定药品不分批库存(以门诊或住院单位)
'参数：int范围=1-门诊,2-住院(缺省)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    strTmp = IIf(int范围 = 1, "门诊", "住院")
    
    '获取药品库存(不分批或分批药品),药房不分批药品不管效期
    strSQL = _
        " Select Nvl(Sum(A.可用数量),0)/Nvl(B." & strTmp & "包装,1) as 库存" & _
        " From 药品库存 A,药品规格 B" & _
        " Where A.药品ID=B.药品ID(+) And A.性质=1" & _
        " And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
        " And A.药品ID=[1] And A.库房ID=[2]" & _
        " Group by Nvl(B." & strTmp & "包装,1)"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng药品ID, lng库房ID)
    If Not rsTmp.EOF Then
        GetStock = Format(rsTmp!库存, "0.00000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetItemField(ByVal lng项目ID As Long, ByVal strField As String) As Variant
'功能：获取指定诊疗项目的指定字段信息
'说明：未处理NULL值
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select " & strField & " From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
    If Not rsTmp.EOF Then GetItemField = rsTmp.Fields(strField).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetGroupCount(ByVal lng组合ID As Long, ByVal int来源 As Integer, Optional bln期效 As Boolean = True) As Long
'功能：获取组合项目中的项目数
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Count(*) as NUM" & _
        " From 诊疗项目组合 A,诊疗项目目录 B,收费项目目录 C" & _
        " Where A.诊疗项目ID=B.ID And A.收费细目ID=C.ID(+) And A.诊疗组合ID=[1]" & _
        " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is NULL) And B.服务对象 IN([2],3)" & _
        " And (A.收费细目ID is NULL Or (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL) And C.服务对象 IN([2],3))" & _
        IIf(bln期效 And int来源 = 1, " And A.期效=1", "")
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng组合ID, int来源)
    If Not rsTmp.EOF Then GetGroupCount = Nvl(rsTmp!Num, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetGroupNone(ByVal lng配方ID As Long, ByVal int来源 As Integer) As String
'功能：读取指定配方中无效的组成中药提示
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    strSQL = "Select B.名称" & _
        " From 诊疗项目组合 A,诊疗项目目录 B,药品规格 C,收费项目目录 D" & _
        " Where A.诊疗项目ID=B.ID And B.ID=C.药名ID And C.药品ID=D.ID And A.诊疗组合ID=[1]" & _
        " And (Not (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is NULL)" & _
        " Or Nvl(B.服务对象,0) Not IN([2],3))" & _
        " And (Not (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 is NULL)" & _
        " Or Nvl(D.服务对象,0) Not IN([2],3))"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng配方ID, int来源)
    Do While Not rsTmp.EOF
        strMsg = strMsg & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    GetGroupNone = Mid(strMsg, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatientFileList(lngDeptID As Long, iFileType As Integer) As ADODB.Recordset
'功能：根据科室和病历文件类型获取可使用的病历文件清单
'参数说明：
'   lngDeptID：科室ID
'   iFileType：文件类型。0-门诊病历;1-住院病历;2-护理记录;3-诊疗文件;4-诊疗单据
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If iFileType = 3 Then
        strSQL = "Select * From 病历文件目录 Where 种类=[1] And 应用<>0 Order by 编号"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1)
    Else
        strSQL = _
            " Select * From 病历文件目录 Where 种类=[1] And " & _
            IIf(lngDeptID = -1, "应用=1", "应用=2 And ','||科室ID||',' Like [2]") & _
            " Order by 编号"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1, "%," & lngDeptID & ",%")
        If rsTmp.EOF Then  '指定科室无该类病历，则查公用病历
            strSQL = "Select * From 病历文件目录 Where 种类=[1] And 应用=1 Order by 编号"
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1)
        End If
    End If
    Set GetPatientFileList = rsTmp
End Function

'将病人本次就诊的文件归档
Public Function PigePatiFile(ByVal lngPatientID As Long, ByVal vPageID As Variant) As Boolean
'lngPatientID：病人ID
'vPageID：挂号单（String）或主页ID（Long）
    PigePatiFile = False
    On Error GoTo DBError
    PigePatiFile_Proc lngPatientID, vPageID
    On Error GoTo 0
    PigePatiFile = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Sub PigePatiFile_Proc(ByVal lngPatientID As Long, ByVal vPageID As Variant)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo DBError
    
    gcnOracle.BeginTrans
    
    If TypeName(vPageID) = "String" Then
        strSQL = "Select ID From 病人病历记录 Where 病人ID=[1] And 挂号单=[2] And 归档日期 Is Null"
    Else
        strSQL = "Select ID From 病人病历记录 Where 病人ID=[1] And 主页ID=[2] And 归档日期 Is Null"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngPatientID, vPageID)
    
    Do While Not rsTmp.EOF
        strSQL = "ZL_病人病历_归档(" & rsTmp(0) & ",'" + UserInfo.姓名 + "')"
        ExecuteProc strSQL, "ZL_病人病历_归档"
'        zlDatabase.ExecuteProcedure "ZL_病人病历_归档(" & rsTmp(0) & ",'" + UserInfo.姓名 + "')", ""
        rsTmp.MoveNext
    Loop

    gcnOracle.CommitTrans
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "病历文件归档"
End Sub

Public Function CalcDrugPrice(ByVal lng药品ID As Long, lng药房ID As Long, ByVal dbl数量 As Double, _
    Optional ByVal str费别 As String, Optional ByVal blnNone加班加价 As Boolean) As Double
'功能：计算药品实价(即然要计算实价,药品则肯定为变价)
'参数：dbl数量=售价数量,按费别打折时计算的是实收金额
'      str费别=是否按费别计算打折的价格,主要在直接计算药品的金额而不显示单价时用
'      gbln加班加价=发送时计算才有用,其它地方都为False
'      blnNone加班加价=为真时,"gbln加班加价"无效
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dbl总数量 As Double, dbl当前数量 As Double
    Dim dbl总金额 As Double, dbl时价 As Double
        
    If dbl数量 = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = _
        " Select Nvl(批次,0) as 批次,Nvl(可用数量,0) as 库存," & _
        " Nvl(Decode(Nvl(实际数量,0),0,0,实际金额/实际数量),0) as 时价" & _
        " From 药品库存" & _
        " Where 库房ID=[1] And 药品ID=[2]" & _
        " And 性质=1 And (Nvl(批次,0)=0 Or 效期 is NULL Or 效期>Trunc(Sysdate))" & _
        " Order by Nvl(批次,0)"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng药房ID, lng药品ID)
    
    dbl总金额 = 0: dbl总数量 = dbl数量
    For i = 1 To rsTmp.RecordCount
        If dbl总数量 = 0 Then Exit For
        If dbl总数量 <= rsTmp!库存 Then
            dbl当前数量 = dbl总数量
        Else
            dbl当前数量 = rsTmp!库存
        End If
        dbl总金额 = dbl总金额 + Format(dbl当前数量 * Format(rsTmp!时价, "0.00000"), gstrDec)
        dbl总数量 = Val(dbl总数量) - Val(dbl当前数量)
        rsTmp.MoveNext
    Next
    If dbl总数量 <> 0 Then
        dbl时价 = 0 '库存不够
    Else
        dbl时价 = Format(dbl总金额 / dbl数量, "0.00000")
        
        '加班加价处理
        If gbln加班加价 And Not blnNone加班加价 Then
            strSQL = _
                " Select To_Number([2])" & _
                    IIf(gbln加班加价, "*Decode(A.加班加价,1,1+Nvl(B.加班加价率,0)/100,1)", "") & " as 金额" & _
                " From 收费项目目录 A,收费价目 B" & _
                " Where A.ID=B.收费细目ID And A.ID=[1]" & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))"
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng药品ID, dbl时价)
            If Not rsTmp.EOF Then dbl时价 = Nvl(rsTmp!金额, 0)
        End If
        
        If str费别 <> "" Then
            '变价项目只有一个收入项目
            strSQL = _
                "Select To_Number([1])*B.实收比率/100 as 金额" & _
                " From 收费价目 A,费别明细 B" & _
                " Where A.收入项目ID=B.收入项目ID And B.费别=[2]" & _
                " And [3] Between B.应收段首值 And B.应收段尾值" & _
                " And A.收费细目ID=[4]"
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", Val(Format(dbl时价 * dbl数量, gstrDec)), str费别, Abs(dbl时价), lng药品ID)
            If Not rsTmp.EOF Then dbl时价 = Nvl(rsTmp!金额, 0)
        End If
    End If
    CalcDrugPrice = dbl时价
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcPrice(ByVal lng项目ID As Long, Optional ByVal str费别 As String, _
    Optional ByVal dbl数量 As Double, Optional ByVal blnNone加班加价 As Boolean) As Double
'功能：获取收费细目的当前售价价格金额,变价返回0
'参数：str费别=是否按费别计算打折的实收金额
'      dbl数量=按费别计算时,必须要传入数量(按售价单位),这时计算的是实收金额
'      gbln加班加价=发送时计算才有用,其它地方都为False
'      blnNone加班加价=为真时,"gbln加班加价"无效
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dbl金额 As Double
    
    On Error GoTo errH
    
    If str费别 = "" Then
        strSQL = _
            " Select Sum(Decode(Nvl(A.是否变价,0),1,NULL," & _
                "B.现价" & IIf(gbln加班加价 And Not blnNone加班加价, "*Decode(A.加班加价,1,1+Nvl(B.加班加价率,0)/100,1)", "") & ")) as 金额" & _
            " From 收费项目目录 A,收费价目 B" & _
            " Where A.ID=B.收费细目ID And A.ID=[1]" & _
            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
        If Not rsTmp.EOF Then dbl金额 = Nvl(rsTmp!金额, 0)
    Else
        '本来可以将ActualMoney函数的SQL一起写在这里，但费别可能被删除而求不出数据
        strSQL = _
            " Select B.收入项目ID,Decode(Nvl(A.是否变价,0),1,NULL," & _
                "B.现价" & IIf(gbln加班加价 And Not blnNone加班加价, "*Decode(A.加班加价,1,1+Nvl(B.加班加价率,0)/100,1)", "") & ") as 金额" & _
            " From 收费项目目录 A,收费价目 B" & _
            " Where A.ID=B.收费细目ID And A.ID=[1]" & _
            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
        For i = 1 To rsTmp.RecordCount
            dbl金额 = dbl金额 + ActualMoney(str费别, rsTmp!收入项目ID, Format(dbl数量 * Format(Nvl(rsTmp!金额, 0), "0.00000"), gstrDec))
            rsTmp.MoveNext
        Next
    End If
    CalcPrice = dbl金额
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ActualMoney(费别 As String, 收入项目ID As Long, 金额 As Currency) As Currency
'功能：根据费别,收入项目ID,金额,求打折后的金额
'说明：金额折扣范围取绝对值范围
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    ActualMoney = 金额
    If 费别 = "" Or 金额 = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = _
        "Select To_Number([1])*实收比率/100 as 金额 From 费别明细" & _
        " Where 收入项目ID=[2] And 费别=[3] And Abs([1]) Between 应收段首值 and 应收段尾值"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", 金额, 收入项目ID, 费别)
    If Not rsTmp.EOF Then ActualMoney = rsTmp!金额
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get临床科室(ByVal int范围 As Integer, Optional ByVal lng病人科室ID As Long, _
    Optional lng缺省科室ID As Long, Optional objCbo As Object, Optional ByVal blnBed As Boolean) As Boolean
'功能：返回临床科室清单或缺省临床科室
'参数：int范围=1-门诊,2-住院,3-门诊或住院
'      lng病人科室ID=病人当前的科室,可能要排开该科室
'      objCbo=要加入科室清单的下拉框,不传时,返回缺省科室
'      lng缺省科室ID=有objCbo时,为缺省定位的科室；否则为要返回的缺省科室
'      blnBed=是否只取有床位的科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
        
    On Error GoTo errH
    
    If int范围 = 1 Then
        strTmp = "1,3"
    ElseIf int范围 = 2 Then
        strTmp = "2,3"
    ElseIf int范围 = 3 Then
        strTmp = "1,2,3"
    End If
    
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,A.简码 " & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And A.ID=B.部门ID And Instr([1],','||B.服务对象||',')>0 And B.工作性质='临床'" & _
        IIf(lng病人科室ID <> 0, " And A.ID<>[2]", "") & _
        IIf(blnBed, " And Exists(Select 科室ID From 床位状况记录 Where 科室ID=A.ID)", "") & _
        " Order by A.编码"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", "," & strTmp & ",", lng病人科室ID)
    
    If Not objCbo Is Nothing Then
        objCbo.Clear
        For i = 1 To rsTmp.RecordCount
            objCbo.AddItem rsTmp!编码 & "-" & rsTmp!名称
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
            If rsTmp!ID = lng缺省科室ID Then
                Call zlControl.CboSetIndex(objCbo.Hwnd, objCbo.NewIndex)
            End If
            rsTmp.MoveNext
        Next
    ElseIf Not rsTmp.EOF Then
        lng缺省科室ID = rsTmp!ID
    End If
    Get临床科室 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Have人员性质(str性质 As String) As Boolean
'功能：判断当前登录人员是否具有指定的人员性质
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    strSQL = "Select A.ID" & _
        " From 人员表 A,人员性质说明 B,上机人员表 C" & _
        " Where A.ID=B.人员ID And B.人员性质=[1]" & _
        " And A.ID=C.人员ID And Upper(C.用户名)=Upper(User)" & _
        " And Rownum=1"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str性质)
    Have人员性质 = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get诊疗项目ID(ByVal lng医嘱ID As Long, bln中药检验 As Boolean) As Long
'功能：读取指定医嘱的诊疗项目ID
'参数：lng医嘱ID=相关ID为NULL的医嘱ID(成药,检查项目,主要手术,中药用法,及独立医嘱)
'      bln中药检验=医嘱ID是否为中药用法或采集方法的ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bln中药检验 Then
        strSQL = "Select 序号,诊疗项目ID From 病人医嘱记录 Where 诊疗类别 IN('7','C') And 相关ID=[1]"
    Else
        strSQL = "Select 序号,诊疗项目ID From 病人医嘱记录 Where ID=[1]"
    End If
    strSQL = strSQL & " Union ALL " & Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    strSQL = strSQL & " Order by 序号"
    
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID)
    If Not rsTmp.EOF Then Get诊疗项目ID = Nvl(rsTmp!诊疗项目ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcAdvicePrice(ByVal lng医嘱ID As Long, Optional ByVal str费别 As String, _
    Optional ByVal bln附加手术 As Boolean, Optional ByVal dbl数量 As Double) As Double
'功能：计算指定非药品医嘱要发送的总金额,以最新价格计算
'参数：str费别=是否按费别计算打折的金额
'      dbl数量=发送的数量,按费别计算时需要(按售价单位),这时计算的是实收金额
'      bln附加手术=是否附加手术医嘱的计价
'      gbln加班加价=发送时计算才有用,其它地方都为False
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dbl单价 As Double, dbl应收 As Double, dbl实收 As Double
    Dim dbl金额 As Double, lng执行科室ID As Long
    Dim lng主收入ID As Long, blnHaveSub As Boolean
    
    On Error GoTo errH
    
    If str费别 = "" And dbl数量 = 0 Then dbl数量 = 1 '忽略发送数量,计算医嘱单价
    
    strSQL = _
        " Select M.主页ID,M.病人科室ID,M.执行科室ID,C.ID,C.类别,C.是否变价,D.跟踪在用,B.收入项目ID,A.数量," & _
        " Nvl(A.从项,0) as 从项,C.加班加价,B.加班加价率,B.附术收费率,Decode(Nvl(C.是否变价,0),1,A.单价,B.现价) as 单价" & _
        " From 病人医嘱记录 M,病人医嘱计价 A,收费价目 B,收费项目目录 C,材料特性 D" & _
        " Where A.收费细目ID=B.收费细目ID And B.收费细目ID=C.ID" & _
        " And C.ID=D.材料ID(+) And M.ID=A.医嘱ID And M.ID=[1]" & _
        " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
        " Order by 从项"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID)
    If gbln从项汇总折扣 And Not rsTmp.EOF And str费别 <> "" Then
        rsTmp.Filter = "从项=1"
        If Not rsTmp.EOF Then blnHaveSub = True
        rsTmp.Filter = 0
    End If
    For i = 1 To rsTmp.RecordCount
        If InStr(",5,6,7,", rsTmp!类别) > 0 And Nvl(rsTmp!是否变价, 0) = 1 Then
            '设定的计价中时价药品单价计算
            lng执行科室ID = Get收费执行科室ID(rsTmp!类别, rsTmp!ID, 4, Nvl(rsTmp!病人科室ID, 0), IIf(Not IsNull(rsTmp!主页ID), 2, 1))
            dbl单价 = Format(CalcDrugPrice(rsTmp!ID, lng执行科室ID, dbl数量 * Nvl(rsTmp!数量, 0), , True), "0.00000")
        ElseIf rsTmp!类别 = "4" And Nvl(rsTmp!是否变价, 0) = 1 And Nvl(rsTmp!跟踪在用, 0) = 1 Then
            '设定的计价中时价卫材单价计算
            lng执行科室ID = Get收费执行科室ID(rsTmp!类别, rsTmp!ID, 4, Nvl(rsTmp!病人科室ID, 0), IIf(Not IsNull(rsTmp!主页ID), 2, 1), Nvl(rsTmp!执行科室ID, 0))
            dbl单价 = Format(CalcDrugPrice(rsTmp!ID, lng执行科室ID, dbl数量 * Nvl(rsTmp!数量, 0), , True), "0.00000")
        Else
            dbl单价 = Format(Nvl(rsTmp!单价, 0), "0.00000")
        End If
        
        '计算应收金额
        dbl应收 = dbl数量 * Nvl(rsTmp!数量, 0) * dbl单价
        If bln附加手术 Then
            dbl应收 = dbl应收 * Nvl(rsTmp!附术收费率, 100) / 100
        End If
        If gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1 Then
            dbl应收 = dbl应收 * (1 + Nvl(rsTmp!加班加价率, 0) / 100)
        End If
        
        '计算实收金额
        If str费别 = "" Then
            dbl应收 = Format(dbl应收, "0.00000")
            dbl实收 = dbl应收
        Else
            dbl应收 = Format(dbl应收, gstrDec)
        
            If gbln从项汇总折扣 And blnHaveSub Then
                If rsTmp!从项 = 0 And lng主收入ID = 0 Then lng主收入ID = rsTmp!收入项目ID
                dbl实收 = dbl应收
            Else
                dbl实收 = Format(ActualMoney(str费别, rsTmp!收入项目ID, CCur(dbl应收)), gstrDec)
            End If
        End If
        
        dbl金额 = dbl金额 + dbl实收
        
        rsTmp.MoveNext
    Next
    
    '套餐整体金额打折
    If gbln从项汇总折扣 And blnHaveSub And lng主收入ID <> 0 And str费别 <> "" Then
        dbl金额 = Format(ActualMoney(str费别, lng主收入ID, CCur(dbl金额)), gstrDec)
    End If
    
    CalcAdvicePrice = dbl金额
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDrugInfo(lng药名ID As Long, lng药品ID As Long, lng药房ID As Long, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal bln停用 As Boolean = True) As ADODB.Recordset
'功能：获取指定药品相关信息
'参数：int范围=1-门诊,2-住院(缺省)
'      bln停用=是否读取已停用药品,用于长嘱药品发送处理
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    strTmp = IIf(int范围 = 1, "门诊", "住院")
    
    strSQL = _
        " Select 药品ID,Sum(Nvl(可用数量,0)) as 库存 From 药品库存" & _
        " Where (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate))" & _
        " And 性质 = 1 And 库房ID=[1]" & IIf(lng药品ID <> 0, " And 药品ID=[2]", "") & _
        " Group by 药品ID Having Sum(Nvl(可用数量,0))<>0"
    strSQL = "Select A.药品ID,A.剂量系数,A." & strTmp & "包装,A." & strTmp & "单位,A.可否分零," & _
        " A.药房分批,B.是否变价,C.库存/A." & strTmp & "包装 as 库存,B.编码,Nvl(D.名称,B.名称) as 名称,B.规格,B.产地,B.撤档时间,B.服务对象" & _
        " From 药品规格 A,收费项目目录 B,(" & strSQL & ") C,收费项目别名 D" & _
        " Where A.药品ID=B.ID And A.药品ID=C.药品ID(+)" & _
        " And B.ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=[5]" & _
        IIf(bln停用, " And B.服务对象 IN([3],3) And (B.撤档时间 is NULL Or B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))", "") & _
        " And A.药名ID=[4]" & IIf(lng药品ID <> 0, " And A.药品ID=[2]", "") & _
        " Order by B.编码"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng药房ID, lng药品ID, int范围, lng药名ID, IIf(gbln商品名, 3, 1))
    Set GetDrugInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function NextNo(intBillID As Integer) As Variant
'功能：根据特定规则产生新的号码,规则如下：
'   一、项目序号：
'   1   病人ID         数字
'   2   住院号         数字
'   3   门诊号         数字
'   10  医嘱发送号     数字,顺序递增编号
'   x   其它单据号     字符,根据编号规则顺序递增编号,不自动补缺
'   二、年度位确定原则:
'       以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码

    Dim rsCtrl As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim vntNo As Variant, strSQL As String
    Dim intYear, strYear As String
    Dim curDate As Date, blnByDate As Boolean
ReStart:
    Err = 0
    On Error GoTo errHand

    If intBillID = 1 Then '病人ID
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From 号码控制表 Where 项目序号=" & intBillID
                Call SQLTest(App.ProductName, "NextNo", strSQL)
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            
            strSQL = "Select Nvl(Max(病人ID),0)+1 as 病人ID From 病人信息 Where 病人ID>=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
            With rsTmp
                'If .State = adStateOpen Then .Close
                'Call SQLTest(App.ProductName, "NextNo", strSQL)
                '.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                'Call SQLTest
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 2 Then '住院号
        '顺序编号还是日期编号
        strSQL = "Select A.*,Sysdate as 日期 From 系统参数表 A Where A.参数号=27"
        With rsTmp
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "NextNo", strSQL)
            .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            Call SQLTest
            If Not .EOF Then
                blnByDate = (IIf(IsNull(!参数值), 1, !参数值) = 2)
                curDate = !日期
            End If
        End With
        
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From 号码控制表 Where 项目序号=" & intBillID
                Call SQLTest(App.ProductName, "NextNo", strSQL)
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(住院号),0)+1 as 住院号 From 病人信息 Where 住院号>=[1]"
            Else
                strSQL = "Select Nvl(Max(住院号),To_Number(To_Char(Sysdate,'YYMM')||'0000'))+1 as 住院号" & _
                    " From 病人信息 Where 住院号 Like To_Number(To_Char(Sysdate,'YYMM'))||'%' And 住院号>=[1]"
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
            With rsTmp
                'If .State = adStateOpen Then .Close
                'Call SQLTest(App.ProductName, "NextNo", strSQL)
                '.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                'Call SQLTest
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            If Not blnByDate Then
                .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "最大号码", IIf(vntNo - 10 > Val(Format(curDate, "YYMM0000")), vntNo - 10, Val(Format(curDate, "YYMM0001")))
            End If
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 3 Then '门诊号
        '顺序编号还是日期编号
        strSQL = "Select A.*,Sysdate as 日期 From 系统参数表 A Where A.参数号=46"
        With rsTmp
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "NextNo", strSQL)
            .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            Call SQLTest
            If Not .EOF Then
                blnByDate = (IIf(IsNull(!参数值), 1, !参数值) = 2)
                curDate = !日期
            End If
        End With
    
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From 号码控制表 Where 项目序号=" & intBillID
                Call SQLTest(App.ProductName, "NextNo", strSQL)
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(门诊号),0)+1 as 门诊号 From 病人信息 Where 门诊号>=[1]"
            Else
                strSQL = "Select Nvl(Max(门诊号),To_Number(To_Char(Sysdate,'YYMMDD')||'0000'))+1 as 门诊号" & _
                    " From 病人信息 Where 门诊号 Like To_Number(To_Char(Sysdate,'YYMMDD'))||'%' And 门诊号>=[1]"
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
            With rsTmp
                'If .State = adStateOpen Then .Close
                'Call SQLTest(App.ProductName, "NextNo", strSQL)
                '.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                'Call SQLTest
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            If Not blnByDate Then
                .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "最大号码", IIf(vntNo - 10 > Val(Format(curDate, "YYMMDD0000")), vntNo - 10, Val(Format(curDate, "YYMMDD0001")))
            End If
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 10 Then '医嘱发送号
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From 号码控制表 C Where C.项目序号=" & intBillID
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "NextNo", strSQL)
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            vntNo = Val(IIf(IsNull(!最大号码), 0, !最大号码)) + 1
            
            On Error Resume Next
            .Update "最大号码", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    Else
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From 号码控制表 C Where C.项目序号=" & intBillID
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "NextNo", strSQL)
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            intYear = Format(!Today, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            vntNo = IIf(IsNull(!最大号码), "", !最大号码)
            
            If IIf(IsNull(!编号规则), 0, !编号规则) = 1 Then
                '按日顺序编号
                If vntNo < strYear & Format(CDate("1992-" & Format(!Today, "MM-dd")) - CDate("1992-01-01") + 1, "000") & "0000" Then
                    vntNo = strYear & Format(CDate("1992-" & Format(!Today, "MM-dd")) - CDate("1992-01-01") + 1, "000") & "0000"
                End If
                vntNo = Left(vntNo, 4) & Right(String(4, "0") & CStr(Val(Mid(vntNo, 5)) + 1), 4)
            Else
                '按年顺序编号
                If Left(vntNo, 1) < strYear Then
                    vntNo = strYear & "0000000"
                End If
                vntNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + 1), 7)
            End If
            
            If Not (UCase(strYear) >= "A" And UCase(strYear) <= "Z") Or zlCommFun.ActualLen(vntNo) > 8 Then GoTo ReStart
            
            On Error Resume Next
            .Update "最大号码", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    NextNo = Null
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'功能：判断是否存在指定处方单据类型的入出类别
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 类别ID From 药品单据性质 Where 单据=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!类别ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDayMoney(lng病人id As Long) As Currency
'功能：获取指定病人当天发生的费用总额
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng病人id)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!金额, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetClinicBillID(ByVal lng项目ID As Long, ByVal int场合 As Integer) As Long
'功能：获取诊疗项目对应的诊疗单据(不管附项,用于生成发送NO)
'参数：int场合=1-门诊,2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 病历文件ID From 诊疗单据应用 Where 诊疗项目ID=[1] And 应用场合=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID, int场合)
    If Not rsTmp.EOF Then GetClinicBillID = Nvl(rsTmp!病历文件ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DeptIsWoman(ByVal lng科室ID As Long) As Boolean
'功能：判断指定科室是否产科
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select * From 部门性质说明 Where 工作性质='产科' And 部门ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng科室ID)
    DeptIsWoman = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitExe(ByVal lng病人id As Long, ByVal lng主页ID As Long) As String
'功能：检查病人在医技科室是否还有未执行完成(未执行或正在执行)的项目
'返回：医技科室名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    strSQL = _
        " Select Distinct A.ID From 部门表 A,部门性质说明 B" & _
        " Where B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 IN('检查','检验','手术','治疗','营养')" & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"
    strSQL = "Select C.名称 as 项目,D.名称 as 科室,B.执行状态" & _
        " From 病人医嘱记录 A,病人医嘱发送 B,诊疗项目目录 C,部门表 D" & _
        " Where A.病人ID=[1] And Nvl(A.主页ID,0)=[2]" & _
        " And B.医嘱ID=A.ID And B.执行部门ID+0 IN(" & strSQL & ")" & _
        " And B.执行状态 IN(0,3) And A.诊疗项目ID=C.ID And B.执行部门ID=D.ID" & _
        " And Not (A.诊疗类别 IN('F','G','D') And A.相关ID is Not NULL)" & _
        " And Not (A.诊疗类别='Z' And Nvl(C.操作类型,'0')<>'0')"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng病人id, lng主页ID)
    strSQL = ""
    For i = 1 To rsTmp.RecordCount
        If i > 10 Then
            strSQL = strSQL & vbCrLf & "... ..."
            Exit For
        Else
            strSQL = strSQL & vbCrLf & rsTmp!项目 & "：在" & rsTmp!科室 & Decode(Nvl(rsTmp!执行状态, 0), 0, "未执行", 3, "正在执行")
        End If
        rsTmp.MoveNext
    Next
    ExistWaitExe = Mid(strSQL, 3)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitDrug(ByVal lng病人id As Long, ByVal lng主页ID As Long) As String
'功能：检查病人在药房是否还有未发药的药品
'返回：药房名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    '以药品收发记录中存在未发药品为准
    strSQL = "Select Distinct C.名称 as 药房" & _
        " From 病人费用记录 A,药品收发记录 B,部门表 C" & _
        " Where A.NO=B.NO And B.库房ID+0=C.ID(+) And A.收费类别 IN('5','6','7')" & _
        " And B.单据 IN(9,10) And Mod(B.记录状态,3)=1 And B.审核人 IS NULL" & _
        " And A.病人ID=[1] And Nvl(A.主页ID,0)=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng病人id, lng主页ID)
    strSQL = ""
    For i = 1 To rsTmp.RecordCount
        strSQL = strSQL & "," & Nvl(rsTmp!药房, "[未定药房]")
        rsTmp.MoveNext
    Next
    ExistWaitDrug = Mid(strSQL, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdvicePause(ByVal lng医嘱ID As Long) As String
'功能：获取指定医嘱的暂停时间段记录
'返回："暂停时间,开始时间;...."
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    strSQL = "Select 操作类型,操作时间 From 病人医嘱状态" & _
        " Where 操作类型 IN(6,7) And 医嘱ID=[1]" & _
        " Order by 操作时间"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!操作类型 = 6 Then
            strTmp = strTmp & ";" & Format(rsTmp!操作时间, "yyyy-MM-dd HH:mm:ss") & ","
        ElseIf rsTmp!操作类型 = 7 Then
            '启用的那一秒不在暂停的范围之内
            strTmp = strTmp & Format(DateAdd("s", -1, rsTmp!操作时间), "yyyy-MM-dd HH:mm:ss")
        End If
        rsTmp.MoveNext
    Next
    GetAdvicePause = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub PrintDiagReport(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, objParent As Object, Optional ByVal PrtMode As Integer = 2, Optional objPic As Object = Nothing, Optional blnMoved As Boolean = False)
'打印辅诊报告
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsImages As ADODB.Recordset
    Dim strRptName As String
    Dim aImages(1, 8) As Variant, aFlagImages(1, 8) As Variant, i As Integer
    Dim strTempPath As String, lngBuffSize As Long
    Dim intReportFormatItem As Integer
    
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strNO As String, lng记录性质 As Long
    Dim iTmpFileCount As Integer, iFlagCount As Integer
    Dim objImages As New DicomImages, intRows As Integer, intCols As Integer, objAssembleImage As New DicomImage
    
    On Error GoTo DBError
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    strTempPath = Mid(strTempPath, 1, InStrRev(strTempPath, "\"))
    
    strSQL = "Select A.NO,A.记录性质,'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2'" & _
        " From 病人医嘱发送 A,病人病历记录 B,病历文件目录 C" & _
        " Where A.报告ID=B.ID And B.文件ID=C.ID And A.医嘱ID=[1] And A.发送号=[2]"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, lng发送号)
    If rsTmp.EOF Then
        MsgBox "该项检查未填写报告，不能打印！", vbInformation, gstrSysName
    Else
        strRptName = rsTmp(2): strNO = rsTmp(0): lng记录性质 = rsTmp(1)
        If PrtMode = 2 Then
            If Not ReportPrintSet(gcnOracle, glngSys, strRptName, objParent) Then Exit Sub
            intReportFormatItem = GetSetting("ZLSOFT", "私有模块\ZLHIS\zl9Report\LocalSet\" & strRptName, "Format", 1)
        Else
            intReportFormatItem = GetSetting("ZLSOFT", "私有模块\ZLHIS\zl9Report\frmReport" & strRptName, "格式", 1)
        End If
        
        'PACS的影像图片
        strSQL = "Select A.用户名1,A.密码1,A.Host1,A.Root1,A.URL1,A.用户名2,A.密码2,A.Host2,A.Root2,A.URL2," & _
            "a.设备号1,a.设备号2,A.NO,A.记录性质 From" & _
            " (Select E.IP地址 As Host1,'/'||E.Ftp目录||'/' as Root1,e.设备号 as 设备号1," & _
            "Decode(D.接收日期,Null,'',to_Char(D.接收日期,'YYYYMMDD')||'/')" & _
            "||D.检查UID||'/'||A.图象文件 As URL1," & _
            "F.IP地址 As Host2,'/'||f.Ftp目录||'/' as Root2," & _
            "Decode(D.接收日期,Null,'',to_Char(D.接收日期,'YYYYMMDD')||'/')" & _
            "||D.检查UID||'/'||A.图象文件 As URL2,f.设备号 as 设备号2," & _
            "C.NO,C.记录性质,E.用户名 as 用户名1,E.密码 as 密码1,F.用户名 as 用户名2,F.密码 as 密码2, Rownum As Seq " & _
            " From 病人病历外部图 A,病人病历内容 B,病人医嘱发送 C,影像检查记录 D,影像设备目录 E,影像设备目录 F" & _
            " Where A.病历ID=B.ID And B.病历记录ID=C.报告ID And C.医嘱ID=D.医嘱ID" & _
            " And C.发送号=D.发送号 And D.位置一=E.设备号(+) and d.位置二=F.设备号(+)" & _
            " And C.医嘱ID=[1] And C.发送号=[2]" & _
            " Order By A.序号) A"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人病历外部图", "H病人病历外部图")
            strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "检查报告", lng医嘱ID, lng发送号, intReportFormatItem)
        iTmpFileCount = rsTmp.RecordCount
        strSQL = "Select A.编号,B.名称,B.W,B.H" & _
            " From zlReports A,zlRPTItems B,病历文件目录 C,病人医嘱记录 D,诊疗单据应用 E" & _
            " Where A.ID=B.报表ID And A.编号='ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2'" & _
            " And C.ID=E.病历文件ID And D.诊疗项目ID=E.诊疗项目ID And Nvl(B.下线,0)=1 And B.类型=11" & _
            " And E.应用场合=D.病人来源 And D.ID=[1]" & _
            " And B.名称 Not Like '标记%' and b.格式号=[3]" & _
            " Order BY b.名称"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人病历外部图", "H病人病历外部图")
            strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rsImages = OpenSQLRecord(strSQL, "检查报告", lng医嘱ID, lng发送号, intReportFormatItem)
        If rsImages.RecordCount = 1 Then
            '图像排版
            For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                
                strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & objFileSystem.GetParentFolderName(rsTmp("URL1"))
                strTmpFile = Replace(strTmpFile, "/", "\")
                MkLocalDir strTmpFile
                strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(rsTmp("URL1"))
                
                If Dir(strTmpFile, vbDirectory) = "" Then
                    If strDeviceNO1 <> rsTmp("设备号1") Then
                        strDeviceNO1 = rsTmp("设备号1")
                        Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("用户名1")), Nvl(rsTmp("密码1"))
                    End If
                    
                    If strDeviceNO2 <> rsTmp("设备号2") Then
                        strDeviceNO2 = rsTmp("设备号2")
                        Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("用户名2")), Nvl(rsTmp("密码2"))
                    End If
                    
                    If Inet1.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL1"))) <> 0 Then
                        Call Inet2.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL2")))
                    End If
                End If
'                objAssembleImage.FileImport strTmpFile, "JPEG"
'                objImages.Add objAssembleImage
                
                objImages.AddNew
                objImages(objImages.Count).FileImport strTmpFile, "JPEG"
                
                rsTmp.MoveNext
            Next
            If objImages.Count > 0 Then
                ResizeRegion i, rsImages("W"), rsImages("H"), intRows, intCols
                Set objAssembleImage = funAssembleImage(objImages, intRows, intCols, rsImages("H"), rsImages("W"))
                strTmpFile = objFileSystem.GetParentFolderName(strTmpFile) & "\" & objFileSystem.GetTempName
                objAssembleImage.FileExport strTmpFile, "JPEG"
                    
                aImages(0, 0) = rsImages("名称")
                aImages(1, 0) = strTmpFile
            End If
            For i = 1 To 8
                aImages(0, i) = "1"
                aImages(1, i) = "1"
            Next
        Else
            For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                If rsImages.EOF Then Exit For
                
    '            strTmpFile = strTempPath & objFileSystem.GetFileName(rsTmp(3))
                
                strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & objFileSystem.GetParentFolderName(rsTmp("URL1"))
                strTmpFile = Replace(strTmpFile, "/", "\")
                MkLocalDir strTmpFile
                strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(rsTmp("URL1"))
                
                If Dir(strTmpFile, vbDirectory) = "" Then
                    If strDeviceNO1 <> rsTmp("设备号1") Then
                        strDeviceNO1 = rsTmp("设备号1")
                        Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("用户名1")), Nvl(rsTmp("密码1"))
                    End If
                    
                    If strDeviceNO2 <> rsTmp("设备号2") Then
                        strDeviceNO2 = rsTmp("设备号2")
                        Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("用户名2")), Nvl(rsTmp("密码2"))
                    End If
                    
                    'Inet.strIPAddress = Nvl(rsTmp(2)): Inet.strUser = Nvl(rsTmp(6)): Inet.strPsw = Nvl(rsTmp(7))
                    If Inet1.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL1"))) <> 0 Then
    '                    strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
    '                        Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL2")))
                    End If
                End If
                    
                aImages(0, i) = rsImages("名称")
                aImages(1, i) = strTmpFile
                rsImages.MoveNext
                rsTmp.MoveNext
            Next
            For i = rsTmp.RecordCount To 8
                aImages(0, i) = "1"
                aImages(1, i) = "1"
            Next
        End If
        
        If Not objPic Is Nothing Then
            '标记图的生成
            strSQL = "Select B.编号,B.名称,A.元素ID,A.内容ID,B.W,B.H From" & _
                " (Select B.ID As 元素ID,A.ID 内容ID,Rownum As Seq From 病人病历内容 A,病历元素目录 B,病人医嘱发送 C" & _
                " Where C.报告ID=A.病历记录ID AND A.元素编码=B.编码 And" & _
                " C.医嘱ID=[1] And C.发送号=[2] And A.元素类型=3) A," & _
                " (Select A.编号,B.名称,B.W,B.H,Rownum As Seq" & _
                " From zlReports A,zlRPTItems B,病历文件目录 C,病人医嘱记录 D,诊疗单据应用 E" & _
                " Where A.ID=B.报表ID And A.编号='ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2'" & _
                " And C.ID=E.病历文件ID And D.诊疗项目ID=E.诊疗项目ID And Nvl(B.下线,0)=1 And B.类型=11" & _
                " And E.应用场合=D.病人来源 And D.ID=[1]" & _
                " And B.名称 Like '标记%'" & _
                " Order BY Trunc(Y/567),Trunc(X/567)) B Where A.Seq=B.Seq"
            'If rsTmp.State <> adStateClosed Then rsTmp.Close
            If blnMoved Then
                strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, lng发送号)
            iFlagCount = rsTmp.RecordCount
            objPic.Cls
            For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                
                strTmpFile = strTempPath & objFileSystem.GetTempName
                
                '计算容器尺寸
                On Error Resume Next
                Set objPic.Picture = ReadCaseMap(rsTmp(2))
                objPic.Width = objPic.ScaleX(objPic.Picture.Width, vbHimetric, vbTwips): objPic.Height = objPic.ScaleY(objPic.Picture.Height, vbHimetric, vbTwips)
                If objPic.Width / objPic.Height > rsTmp(4) / rsTmp(5) Then
                    objPic.Width = objPic.Height * rsTmp(4) / rsTmp(5)
                Else
                    objPic.Height = objPic.Width / (rsTmp(4) / rsTmp(5))
                End If
                objPic.Cls: Set objPic.Picture = Nothing
                On Error GoTo DBError
                Call ShowMapInOjbect(objPic, rsTmp(2), rsTmp(3), blnMoved:=blnMoved)
                SavePicture objPic.Image, strTmpFile
                objPic.Cls
            
                aFlagImages(0, i) = rsTmp(1)
                aFlagImages(1, i) = strTmpFile
                rsTmp.MoveNext
            Next
            For i = rsTmp.RecordCount To 8
                aFlagImages(0, i) = "1"
                aFlagImages(1, i) = "1"
            Next
            Call ReportOpen(gcnOracle, glngSys, strRptName, objParent, _
                "NO=" & strNO, _
                "性质=" & lng记录性质, "医嘱ID=" & lng医嘱ID, _
                "ReportFormat=" & gintReportFormat, _
                aImages(0, 0) & "=" & aImages(1, 0), _
                aImages(0, 1) & "=" & aImages(1, 1), _
                aImages(0, 2) & "=" & aImages(1, 2), _
                aImages(0, 3) & "=" & aImages(1, 3), _
                aImages(0, 4) & "=" & aImages(1, 4), _
                aImages(0, 5) & "=" & aImages(1, 5), _
                aImages(0, 6) & "=" & aImages(1, 6), _
                aImages(0, 7) & "=" & aImages(1, 7), _
                aImages(0, 8) & "=" & aImages(1, 8), _
                aFlagImages(0, 0) & "=" & aFlagImages(1, 0), _
                aFlagImages(0, 1) & "=" & aFlagImages(1, 1), _
                aFlagImages(0, 2) & "=" & aFlagImages(1, 2), _
                aFlagImages(0, 3) & "=" & aFlagImages(1, 3), _
                aFlagImages(0, 4) & "=" & aFlagImages(1, 4), _
                aFlagImages(0, 5) & "=" & aFlagImages(1, 5), _
                aFlagImages(0, 6) & "=" & aFlagImages(1, 6), _
                aFlagImages(0, 7) & "=" & aFlagImages(1, 7), _
                aFlagImages(0, 8) & "=" & aFlagImages(1, 8), PrtMode)
        Else
            Call ReportOpen(gcnOracle, glngSys, strRptName, objParent, _
                "NO=" & strNO, _
                "性质=" & lng记录性质, "医嘱ID=" & lng医嘱ID, _
                "ReportFormat=" & gintReportFormat, _
                aImages(0, 0) & "=" & aImages(1, 0), _
                aImages(0, 1) & "=" & aImages(1, 1), _
                aImages(0, 2) & "=" & aImages(1, 2), _
                aImages(0, 3) & "=" & aImages(1, 3), _
                aImages(0, 4) & "=" & aImages(1, 4), _
                aImages(0, 5) & "=" & aImages(1, 5), _
                aImages(0, 6) & "=" & aImages(1, 6), _
                aImages(0, 7) & "=" & aImages(1, 7), _
                aImages(0, 8) & "=" & aImages(1, 8), PrtMode)
        End If
        '删除临时文件
'        For i = 0 To iTmpFileCount - 1
'            objFileSystem.DeleteFile aImages(1, i), True
'        Next
'        For i = 0 To iFlagCount - 1
'            objFileSystem.DeleteFile aFlagImages(1, i), True
'        Next
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub PrintDiagRpt_New(ByVal lng报告ID As Long, objParent As Object, Optional ByVal PrtMode As Integer = 2, Optional objPic As Object = Nothing, Optional ByVal blnMoved As Boolean)
'功能：打印辅诊报告
'参数：blnMoved=该病人诊治数据是否已转出
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strRptName As String
    Dim aImages(1, 8) As Variant, aFlagImages(1, 8) As Variant, i As Integer
    Dim strTempPath As String, lngBuffSize As Long
    Dim intReportFormatItem As Integer
    
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strNO As String, lng记录性质 As Long, lng医嘱ID As Long, lng发送号 As Long
    Dim iTmpFileCount As Integer, iFlagCount As Integer
    
    On Error GoTo DBError
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    strTempPath = Mid(strTempPath, 1, InStrRev(strTempPath, "\"))
    
    strSQL = "Select A.医嘱ID,A.发送号,A.NO,A.记录性质,'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2' As 报表编号,D.相关ID,E.类别,E.操作类型" & _
        " From 病人医嘱发送 A,病人病历记录 B,病历文件目录 C,病人医嘱记录 D,诊疗项目目录 E" & _
        " Where A.报告ID=B.ID And B.文件ID=C.ID And A.医嘱ID=D.ID And D.诊疗项目ID=E.ID And A.报告ID=[1] Order By D.相关ID Desc"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng报告ID)
    If rsTmp.EOF Then
        MsgBox "该项检查未填写报告，不能打印！", vbInformation, gstrSysName
    Else
        lng医嘱ID = rsTmp("医嘱ID"): lng发送号 = rsTmp("发送号")
        strRptName = rsTmp("报表编号"): strNO = rsTmp("NO"): lng记录性质 = rsTmp("记录性质")
        If PrtMode = 2 Then
            If Not ReportPrintSet(gcnOracle, glngSys, strRptName, objParent) Then Exit Sub
            intReportFormatItem = GetSetting("ZLSOFT", "私有模块\ZLHIS\zl9Report\LocalSet\" & strRptName, "Format", 1)
        Else
            intReportFormatItem = GetSetting("ZLSOFT", "私有模块\ZLHIS\zl9Report\frmReport" & strRptName, "格式", 1)
        End If
        
        '检验
        If Nvl(rsTmp("类别") = "E") And Nvl(rsTmp("操作类型")) = "6" Then
            strSQL = "Select A.医嘱ID,A.发送号,A.NO,A.记录性质" & _
                " From 病人医嘱发送 A,检验标本记录 C,检验项目分布 D" & _
                " Where D.标本ID+0=C.ID And C.医嘱ID+0=A.医嘱ID And D.医嘱ID=[1]"
            If blnMoved Then
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
                strSQL = Replace(strSQL, "检验标本记录", "H检验标本记录")
                strSQL = Replace(strSQL, "检验项目分布", "H检验项目分布")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID)
            If Not rsTmp.EOF Then
                lng医嘱ID = rsTmp("医嘱ID"): lng发送号 = rsTmp("发送号")
                strNO = rsTmp("NO"): lng记录性质 = rsTmp("记录性质")
            End If
        End If
        
        strSQL = "Select B.编号,B.名称,A.用户名1,A.密码1,A.Host1,A.Root1,A.URL1,A.用户名2,A.密码2,A.Host2,A.Root2,A.URL2," & _
            "a.设备号1,a.设备号2,A.NO,A.记录性质 From" & _
            " (Select E.IP地址 As Host1,'/'||E.Ftp目录||'/' as Root1,e.设备号 as 设备号1," & _
            "Decode(D.接收日期,Null,'',to_Char(D.接收日期,'YYYYMMDD')||'/')" & _
            "||D.检查UID||'/'||A.图象文件 As URL1," & _
            "F.IP地址 As Host2,'/'||f.Ftp目录||'/' as Root2," & _
            "Decode(D.接收日期,Null,'',to_Char(D.接收日期,'YYYYMMDD')||'/')" & _
            "||D.检查UID||'/'||A.图象文件 As URL2,f.设备号 as 设备号2," & _
            "C.NO,C.记录性质,E.用户名 as 用户名1,E.密码 as 密码1,F.用户名 as 用户名2,F.密码 as 密码2, Rownum As Seq " & _
            " From 病人病历外部图 A,病人病历内容 B,病人医嘱发送 C,影像检查记录 D,影像设备目录 E,影像设备目录 F" & _
            " Where A.病历ID=B.ID And B.病历记录ID=C.报告ID And C.医嘱ID=D.医嘱ID" & _
            " And C.发送号=D.发送号 And D.位置一=E.设备号(+) and d.位置二=F.设备号(+)" & _
            " And C.医嘱ID=[1] And C.发送号=[2]" & _
            " Order By A.序号) A," & _
            " (select z.编号,z.名称,rownum as seq " & _
            " from " & _
            " (Select A.编号,B.名称,Rownum As Seq" & _
            " From zlReports A,zlRPTItems B,病历文件目录 C,病人医嘱记录 D,诊疗单据应用 E" & _
            " Where A.ID=B.报表ID And A.编号='ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2'" & _
            " And C.ID=E.病历文件ID And D.诊疗项目ID=E.诊疗项目ID And Nvl(B.下线,0)=1 And B.类型=11" & _
            " And E.应用场合=D.病人来源 And D.ID=[1]" & _
            " And B.名称 Not Like '标记%' and b.格式号=[3]" & _
            " Order BY b.名称 ) z ) B Where A.Seq=B.Seq"
        If blnMoved Then
            strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
            strSQL = Replace(strSQL, "病人病历外部图", "H病人病历外部图")
        End If
        'If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, lng发送号, intReportFormatItem)
        iTmpFileCount = rsTmp.RecordCount
        If PrtMode = 2 Then
            If Not ReportPrintSet(gcnOracle, glngSys, strRptName, objParent) Then Exit Sub
        End If
        For i = 0 To rsTmp.RecordCount - 1
            If i > 8 Then Exit For
            
            strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & objFileSystem.GetParentFolderName(rsTmp("URL1"))
            strTmpFile = Replace(strTmpFile, "/", "\")
            MkLocalDir strTmpFile
            strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(rsTmp("URL1"))
            
            If Dir(strTmpFile, vbDirectory) = "" Then
                If strDeviceNO1 <> rsTmp("设备号1") Then
                    strDeviceNO1 = rsTmp("设备号1")
                    Inet1.FuncFtpConnect rsTmp("Host1"), rsTmp("用户名1"), rsTmp("密码1")
                End If
                
                If strDeviceNO2 <> rsTmp("设备号2") Then
                    strDeviceNO2 = rsTmp("设备号2")
                    Inet2.FuncFtpConnect rsTmp("Host2"), rsTmp("用户名2"), rsTmp("密码2")
                End If
                
                'Inet.strIPAddress = Nvl(rsTmp(2)): Inet.strUser = Nvl(rsTmp(6)): Inet.strPsw = Nvl(rsTmp(7))
                If Inet1.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL1"))) <> 0 Then
'                    strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
'                        Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
                    Call Inet2.FuncDownloadFile(objFileSystem.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL2")))
                End If
            End If
            
            aImages(0, i) = rsTmp(1)
            aImages(1, i) = strTmpFile
            rsTmp.MoveNext
        Next
        For i = rsTmp.RecordCount To 8
            aImages(0, i) = "1"
            aImages(1, i) = "1"
        Next
        If Not objPic Is Nothing Then
            '标记图的生成
            strSQL = "Select B.编号,B.名称,A.元素ID,A.内容ID,B.W,B.H From" & _
                " (Select B.ID As 元素ID,A.ID 内容ID,Rownum As Seq From 病人病历内容 A,病历元素目录 B,病人医嘱发送 C" & _
                " Where C.报告ID=A.病历记录ID AND A.元素编码=B.编码 And" & _
                " C.医嘱ID=[1] And C.发送号=[2] And A.元素类型=3) A," & _
                " (Select A.编号,B.名称,B.W,B.H,Rownum As Seq" & _
                " From zlReports A,zlRPTItems B,病历文件目录 C,病人医嘱记录 D,诊疗单据应用 E" & _
                " Where A.ID=B.报表ID And A.编号='ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2'" & _
                " And C.ID=E.病历文件ID And D.诊疗项目ID=E.诊疗项目ID And Nvl(B.下线,0)=1 And B.类型=11" & _
                " And E.应用场合=D.病人来源 And D.ID=[1]" & _
                " And B.名称 Like '标记%'" & _
                " Order BY Trunc(Y/567),Trunc(X/567)) B Where A.Seq=B.Seq"
            If blnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
                strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
            End If
            'If rsTmp.State <> adStateClosed Then rsTmp.Close
            Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, lng发送号)
            iFlagCount = rsTmp.RecordCount
            objPic.Cls
             For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                
                strTmpFile = strTempPath & objFileSystem.GetTempName
                
                '计算容器尺寸
                On Error Resume Next
                Set objPic.Picture = ReadCaseMap(rsTmp(2))
                objPic.Width = objPic.ScaleX(objPic.Picture.Width, vbHimetric, vbTwips): objPic.Height = objPic.ScaleY(objPic.Picture.Height, vbHimetric, vbTwips)
                If objPic.Width / objPic.Height > rsTmp(4) / rsTmp(5) Then
                    objPic.Width = objPic.Height * rsTmp(4) / rsTmp(5)
                Else
                    objPic.Height = objPic.Width / (rsTmp(4) / rsTmp(5))
                End If
                objPic.Cls: Set objPic.Picture = Nothing
                On Error GoTo DBError
                Call ShowMapInOjbect(objPic, rsTmp(2), rsTmp(3), blnMoved:=blnMoved)
                SavePicture objPic.Image, strTmpFile
                objPic.Cls
            
                aFlagImages(0, i) = rsTmp(1)
                aFlagImages(1, i) = strTmpFile
                rsTmp.MoveNext
            Next
            For i = rsTmp.RecordCount To 8
                aFlagImages(0, i) = "1"
                aFlagImages(1, i) = "1"
            Next
            Call ReportOpen(gcnOracle, glngSys, strRptName, objParent, _
                "NO=" & strNO, _
                "性质=" & lng记录性质, "医嘱ID=" & lng医嘱ID, _
                aImages(0, 0) & "=" & aImages(1, 0), _
                aImages(0, 1) & "=" & aImages(1, 1), _
                aImages(0, 2) & "=" & aImages(1, 2), _
                aImages(0, 3) & "=" & aImages(1, 3), _
                aImages(0, 4) & "=" & aImages(1, 4), _
                aImages(0, 5) & "=" & aImages(1, 5), _
                aImages(0, 6) & "=" & aImages(1, 6), _
                aImages(0, 7) & "=" & aImages(1, 7), _
                aImages(0, 8) & "=" & aImages(1, 8), _
                aFlagImages(0, 0) & "=" & aFlagImages(1, 0), _
                aFlagImages(0, 1) & "=" & aFlagImages(1, 1), _
                aFlagImages(0, 2) & "=" & aFlagImages(1, 2), _
                aFlagImages(0, 3) & "=" & aFlagImages(1, 3), _
                aFlagImages(0, 4) & "=" & aFlagImages(1, 4), _
                aFlagImages(0, 5) & "=" & aFlagImages(1, 5), _
                aFlagImages(0, 6) & "=" & aFlagImages(1, 6), _
                aFlagImages(0, 7) & "=" & aFlagImages(1, 7), _
                aFlagImages(0, 8) & "=" & aFlagImages(1, 8), PrtMode)
        Else
            Call ReportOpen(gcnOracle, glngSys, strRptName, objParent, _
                "NO=" & strNO, _
                "性质=" & lng记录性质, "医嘱ID=" & lng医嘱ID, _
                aImages(0, 0) & "=" & aImages(1, 0), _
                aImages(0, 1) & "=" & aImages(1, 1), _
                aImages(0, 2) & "=" & aImages(1, 2), _
                aImages(0, 3) & "=" & aImages(1, 3), _
                aImages(0, 4) & "=" & aImages(1, 4), _
                aImages(0, 5) & "=" & aImages(1, 5), _
                aImages(0, 6) & "=" & aImages(1, 6), _
                aImages(0, 7) & "=" & aImages(1, 7), _
                aImages(0, 8) & "=" & aImages(1, 8), PrtMode)
        End If
        '删除临时文件
        For i = 0 To iTmpFileCount - 1
            objFileSystem.DeleteFile aImages(1, i), True
        Next
        For i = 0 To iFlagCount - 1
            objFileSystem.DeleteFile aFlagImages(1, i), True
        Next
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function ItemIsVarPrice(ByVal lng项目ID As Long) As Boolean
'功能：判断指定项目是否变价(非药品和跟踪在用的卫材)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select A.类别,A.是否变价,B.跟踪在用 From 收费项目目录 A,材料特性 B Where A.ID=B.材料ID(+) And A.ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
    If Not rsTmp.EOF Then
        If Not (InStr(",5,6,7,", rsTmp!类别) > 0 Or rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1) Then
            ItemIsVarPrice = Nvl(rsTmp!是否变价, 0) <> 0
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlClinicCodeRepeat(str编码 As String, Optional lng项目ID As Long) As Boolean
'功能：检查诊疗项目编码的是否与现有编码重复，重复则给出提示
'入参：str编码-输入的编码；lng项目ID-自己的ID号，当修改时，需要将自身除开才能判断
'出参：重复返回True；否则反馈Flase
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select K.名称||' ['||I.编码||']'||I.名称 as 名称" & _
        " From 诊疗项目目录 I,诊疗项目类别 K" & _
        " Where I.类别=K.编码 And I.编码=[1] And I.ID<>[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str编码, lng项目ID)
    If Not rsTmp.EOF Then
        MsgBox "该项目编码与“" & rsTmp!名称 & "”的编码重复！", vbInformation, gstrSysName
        zlClinicCodeRepeat = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetSymbol(StrInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode('" & StrInput & "') from dual"
    Else
        strSQL = "select zlSpellcode('" & StrInput & "') from dual"
    End If
    On Error GoTo errHand
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function GetSendMoneyState(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, Optional ByVal blnMoved As Boolean) As String
'功能：获取指定医嘱某次发送之后的计费状态，主要考虑一些组合医嘱有多种计费的状态
'参数：lng医嘱ID=检验组合主项目,手术主项目,第一个检验项目的医嘱ID(即在医技站中显示的项目的)
'返回：",-1,0,1,"：其中-1=无需计费,1=已计费,0=未计费
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 相关ID From 病人医嘱记录 Where 诊疗类别='C' And ID=[1]"
    strSQL = _
        " Select ID From 病人医嘱记录 Where ID=[1] Or (相关ID=[1] And 诊疗类别 IN('F','D'))" & _
        " Union ALL " & _
        " Select ID From 病人医嘱记录 Where 诊疗类别='C' And 相关ID=(" & strSQL & ")"
    strSQL = "Select Distinct 计费状态 From 病人医嘱发送 Where 医嘱ID IN(" & strSQL & ") And 发送号=[2]"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, lng发送号)
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & Nvl(rsTmp!计费状态, 0)
        rsTmp.MoveNext
    Loop
    If strSQL <> "" Then GetSendMoneyState = strSQL & ","
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillMax序号(ByVal strNO As String, ByVal int记录性质 As Integer, str登记时间 As String) As Integer
'功能：获取指定单据当前的最大序号+1
'参数：str登记时间=组合医嘱只生成了部份主费用时，将要新生成的收费划价单(NO相同)的时间与已生成的一致。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    str登记时间 = ""
    strSQL = "Select Max(序号) as 序号,Max(登记时间) as 时间 From 病人费用记录 Where NO=[1] And 记录性质=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strNO, int记录性质)
    If Not rsTmp.EOF Then
        GetBillMax序号 = Nvl(rsTmp!序号, 0) + 1
        If Not IsNull(rsTmp!时间) Then
            str登记时间 = Format(rsTmp!时间, "yyyy-MM-dd HH:mm:ss")
        End If
    Else
        GetBillMax序号 = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistBalance(ByVal strNO As String) As Boolean
'功能：判断指定的收费划价单是否存在已经收费的内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ID From 病人费用记录 Where 记录性质=1 And 记录状态 IN(1,3) And NO=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strNO)

    BillExistBalance = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRegistRoom(ByVal str号别 As String, ByVal lng号别ID As Long, ByVal int分诊 As Integer) As String
'功能：根据号别的分诊方式获取号别的诊室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If int分诊 = 0 Then Exit Function '不分诊
    
    On Error GoTo errH
    
    '处理分诊
    If int分诊 = 1 Then
        '指定诊室
        strSQL = "Select 门诊诊室 From 挂号安排诊室 Where 号表ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng号别ID)
        If Not rsTmp.EOF Then GetRegistRoom = rsTmp!门诊诊室
    ElseIf int分诊 = 2 Then
        '动态分诊：该个号别当天挂号未诊数最少的诊室   //todo未考虑预约挂号
        strSQL = _
            " Select 门诊诊室,Sum(NUM) as NUM From (" & _
                " Select 门诊诊室,0 as NUM From 挂号安排诊室 Where 号表ID=[1]" & _
                " Union ALL" & _
                " Select 诊室,Count(诊室) as NUM From 病人挂号记录" & _
                " Where Nvl(执行状态,0)=0" & _
                " And 登记时间 Between Trunc(Sysdate) And Sysdate And 号别=[2]" & _
                " And 诊室 IN(Select 门诊诊室 From 挂号安排诊室 Where 号表ID=[1])" & _
                " Group by 诊室)" & _
            " Group by 门诊诊室 Order by Num"
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng号别ID, str号别)
        If Not rsTmp.EOF Then GetRegistRoom = rsTmp!门诊诊室
    ElseIf int分诊 = 3 Then
        '平均分诊：当前分配=1表示下次应取的当前诊室
        strSQL = "Select * From 挂号安排诊室 Where 号表ID=" & lng号别ID
        Call OpenRecord(rsTmp, strSQL, "mdlPublic", adOpenStatic, adLockOptimistic) '可写记录集
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If Nvl(rsTmp!当前分配, 0) = 1 Then
                    GetRegistRoom = rsTmp!门诊诊室
                    rsTmp!当前分配 = 0
                    
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!当前分配 = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '处理第一次平均分配
            If GetRegistRoom = "" Then
                rsTmp.MoveFirst
                GetRegistRoom = rsTmp!门诊诊室
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!当前分配 = 1
                rsTmp.Update
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadCaseMap(lngID As Long) As StdPicture
'功能：根据标记图ID返回图形对象
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 图形 From 病历标记图 Where 元素ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngID)
        
    If rsTmp.EOF Then Exit Function
    If IsNull(rsTmp!图形) Then Exit Function
    
    On Error GoTo 0
    
    intFile = FreeFile
    strFile = CurDir & "\zlNewPicture" & Timer & ".pic"
    
    Open strFile For Binary As intFile
    
    lngFileSize = rsTmp.Fields("图形").ActualSize
    ReDim arrBin(lngFileSize - 1) As Byte
    arrBin() = rsTmp.Fields("图形").GetChunk(lngFileSize)
    Put intFile, , arrBin()
    Close intFile
    
    Set ReadCaseMap = VB.LoadPicture(strFile)
    Kill strFile
    Exit Function
errH:
    Kill strFile
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DeptExist(ByVal str工作性质 As String, ByVal int服务对象 As Integer) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 部门ID From 部门性质说明 Where 工作性质=[1] And 服务对象 IN([2],3) And Rownum<2"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str工作性质, int服务对象)
    DeptExist = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExpend(ByVal strNO As String) As Boolean
'功能：判断挂号单是否已经超过有效挂号天数。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If gint挂号天数 = 0 Then Exit Function
        
    On Error GoTo errH
    
    '按时点算
    strSQL = "Select Sysdate-登记时间 as 间隔 From 病人挂号记录 Where NO=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strNO)
    If Not rsTmp.EOF Then
        BillExpend = rsTmp!间隔 > gint挂号天数
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedByDate(ByVal vDate As Date) As Boolean
'功能：判断指定日期之前的是否可能已经执行了数据转出
'参数：vDate=时间点或时间段的开始时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 上次日期 From zlDataMove Where 系统=[1] And 组号=1 And 上次日期 is Not Null"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '上次日期没有时点,"<"判断与转出过程中一致
        If vDate < rsTmp!上次日期 Then
            MovedByDate = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedByNO(ByVal strNO As String, ByVal strTable As String, Optional ByVal strWhere As String) As Boolean
'功能：判断指定日期之前的是否可能已经执行了数据转出
'参数：vDate=时间点或时间段的开始时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select NO From H" & strTable & " Where NO=[1] And Rownum<2" & IIf(strWhere <> "", " And " & strWhere, "")
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", strNO)
    If Not rsTmp.EOF Then
        MovedByNO = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedBySend(ByVal lng医嘱ID As Long, Optional ByVal lng发送号 As Long) As Boolean
'功能：检查某次发送的医嘱中的费用是否已经执行了数据转出
'参数：lng发送号=因为门诊医嘱只有一次发送,可以不传入
'说明：1.在医嘱未转出的情况下，执行回退或作废操作时，如果包含已转出的费用，则禁止
'      2.对于住院长嘱有多次发送的情况，只判断当前要回退的这次医嘱发送费用
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1]"
    strSQL = "Select B.NO From 病人医嘱发送 A,H病人费用记录 B" & _
        " Where A.记录性质=B.记录性质 And A.NO=B.NO" & _
        IIf(lng发送号 <> 0, " And A.发送号+0=[2]", "") & _
        " And A.医嘱ID IN(" & strSQL & ")" & _
        " Group by B.NO"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, lng发送号)
    If Not rsTmp.EOF Then MovedBySend = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get医疗付款码(ByVal str名称 As String) As String
'功能：根据医疗付款方式名称获取医疗付款编码
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If str名称 = "" Then Exit Function
    
    strSQL = "Select 编码 From 医疗付款方式 Where 名称=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str名称)
    If Not rsTmp.EOF Then Get医疗付款码 = Nvl(rsTmp!编码)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckOneDuty(ByVal str医嘱 As String, ByVal str职务 As String, ByVal str医生 As String, ByVal bln医保 As Boolean) As String
'功能：检查当前指定药品处方职务是否符合
'参数：str医嘱=药品医嘱提示内容
'      str职务=药品处方职务
'      str医生=开嘱医生
'      bln医保=是否公费或医保病人
'      grsDuty=记录医生职务缓存
'返回：职务不满足的提示信息，如果满足则返回空。
    Const STR_职务 = "正高,副高,中级,助理/师级,员/士,,,,待聘"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim int职务A As Integer, int职务B As Integer
    
    If Len(str职务) <> 2 Or str医生 = "" Then Exit Function
    
    '取药品处方职务
    If bln医保 Then
        int职务B = Val(Right(str职务, 1))
    Else
        int职务B = Val(Left(str职务, 1))
    End If
    If int职务B = 0 Then Exit Function '不限制
    
    '取医生职务
    If grsDuty Is Nothing Then
        Set grsDuty = New ADODB.Recordset
        grsDuty.Fields.Append "医生", adVarChar, 50
        grsDuty.Fields.Append "职务", adInteger
        grsDuty.CursorLocation = adUseClient
        grsDuty.LockType = adLockOptimistic
        grsDuty.CursorType = adOpenStatic
        grsDuty.Open
    End If
    grsDuty.Filter = "医生='" & str医生 & "'"
    If grsDuty.EOF Then
        On Error GoTo errH
        strSQL = "Select 姓名,Nvl(聘任技术职务,0) as 职务 From 人员表 Where 姓名=[1]"
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", str医生)
        On Error GoTo 0
        If Not rsTmp.EOF Then
            grsDuty.AddNew
            grsDuty!医生 = rsTmp!姓名
            grsDuty!职务 = rsTmp!职务
            grsDuty.Update
        End If
    End If
    If Not grsDuty.EOF Then
        int职务A = grsDuty!职务
    End If
        
    '检查职务要求
    If int职务A = 0 Then
        '医生未设置职务的情况
        strMsg = """" & str医嘱 & """要求的处方职务不满足：" & vbCrLf & vbCrLf & IIf(bln医保, "对医保或公费病人,", "") & _
            "该药品要求职务至少为""" & Split(STR_职务, ",")(int职务B - 1) & """才能下达,而医生""" & str医生 & """未设置职务。"
    ElseIf int职务B < int职务A Then
        '数值越小职务越高
        strMsg = """" & str医嘱 & """要求的处方职务不满足：" & vbCrLf & vbCrLf & IIf(bln医保, "对医保或公费病人,", "") & _
            "该药品要求职务至少为""" & Split(STR_职务, ",")(int职务B - 1) & """才能下达,而医生""" & str医生 & """的职务为""" & Split(STR_职务, ",")(int职务A - 1) & """。"
    End If
    CheckOneDuty = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiHaveOut(ByVal lng病人id As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
'功能：判断病人是否已出院(包括预出院),用于并发操作判断
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 病人ID From 病案主页 Where (出院日期 is Not Null Or Nvl(状态,0)=3) And 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng病人id, lng主页ID)
    If Not rsTmp.EOF Then PatiHaveOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlExpense", intNum)
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '按年编号
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'将一个检查的影像文件转移到另外检查中去
Public Function MergeImageFiles(ByVal strCurrUID As String, ByVal strNewUID As String, _
    Optional ByVal strReceiveDate As String = "", Optional ByVal strMoveFiles As String = "") As Boolean
    
    Dim objSrcFtp As New clsFtp, objDestFtp As New clsFtp
    Dim strSrcPath As String, strDestPath As String
    Dim rsTmp As New ADODB.Recordset, strSQL As String, strTmpFile As String
    Dim aFiles() As String, i As Integer, objFile As New Scripting.FileSystemObject
    '存储原检查UID的FTP连接
    Dim strFTPUser As String, strFTPPassw As String, strFTPHost As String, strFTPRoot As String
    On Error GoTo errH
    MergeImageFiles = True
    If strCurrUID = strNewUID Then Exit Function
    
    '初始化源Ftp
    strSQL = "Select D.用户名 As FtpUser,D.密码 As FtpPwd," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL " & _
        "From 影像检查记录 C,影像设备目录 D " & _
        "Where Decode(C.位置二,Null,C.位置一,C.位置二)=D.设备号(+)" & _
        "And C.检查UID= [1] Union All " & _
        "Select D.用户名 As FtpUser,D.密码 As FtpPwd," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL " & _
        "From 影像临时记录 C,影像设备目录 D " & _
        "Where Decode(C.位置二,Null,C.位置一,C.位置二)=D.设备号(+)" & _
        "And C.检查UID= [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "ZLPACSWork", strCurrUID)
    If rsTmp.EOF Then
        MergeImageFiles = False
        Exit Function
    End If
    
    '存储原有FTP连接设置
    strFTPHost = rsTmp("Host")
    strFTPPassw = rsTmp("FtpPwd")
    strFTPRoot = rsTmp("Root")
    strFTPUser = rsTmp("FtpUser")
    
    With objSrcFtp
        '.strIPAddress = Nvl(rsTmp("Host")): .strUser = Nvl(rsTmp("FtpUser")): .strPsw = Nvl(rsTmp("FtpPwd"))
        .FuncFtpConnect rsTmp("Host"), rsTmp("FtpUser"), rsTmp("FtpPwd")
        strSrcPath = rsTmp("Root") & Nvl(rsTmp("URL"))
    End With
    
    '初始化目标Ftp,如果目标UID不存在，创建一个新路径
    Set rsTmp = OpenSQLRecord(strSQL, "ZLPACSWork", strNewUID)
    If rsTmp.EOF Then
        If strReceiveDate <> "" Then
            With objDestFtp
                .FuncFtpConnect strFTPHost, strFTPUser, strFTPPassw
                strDestPath = strFTPRoot & Format(strReceiveDate, "YYYYMMDD") & "/" & strNewUID
                '创建FTP目录
                .FuncFtpMkDir strFTPRoot, Format(strReceiveDate, "YYYYMMDD") & "/" & strNewUID
            End With
        Else
            MergeImageFiles = False
            Exit Function
        End If
    Else
        With objDestFtp
    '        .strIPAddress = Nvl(rsTmp("Host")): .strUser = Nvl(rsTmp("FtpUser")): .strPsw = Nvl(rsTmp("FtpPwd"))
            .FuncFtpConnect rsTmp("Host"), rsTmp("FtpUser"), rsTmp("FtpPwd")
            strDestPath = rsTmp("Root") & Nvl(rsTmp("URL"))
        End With
    End If
    
    '提取需要移动的文件名
    If strMoveFiles <> "" Then
        aFiles = Split(strMoveFiles, "|")
    Else
        aFiles = Split(objSrcFtp.FuncDirFiles(strSrcPath), "|")
    End If
    
    For i = 0 To UBound(aFiles)
        strTmpFile = App.Path & "\" & aFiles(i)
        Call objSrcFtp.FuncDownloadFile(strSrcPath, strTmpFile, aFiles(i))
        Call objDestFtp.FuncUploadFile(strDestPath, strTmpFile, aFiles(i))
        
        Kill strTmpFile
        Call objSrcFtp.FuncDelFile(strSrcPath, aFiles(i))
    Next
    
    '需要测试一下，如果有图象在目录中，目录是否会删除？
    Call objSrcFtp.FuncFtpDelDir(Replace(strSrcPath, strCurrUID, ""), strCurrUID)
    
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    MergeImageFiles = False
    Call SaveErrLog
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'功能：创建本地目录
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function GetAllImageFiles(ByVal CheckUID As String, _
    Optional ByVal strSerials As String = "", Optional ByVal blnMoved As Boolean = False, _
    Optional strFTPHost As String, Optional strDicomPath As String, Optional strLocalPath As String, _
    Optional strFTPUser As String, Optional strFtpPwd As String) As String()
    Dim strSQL As String, lngSeqUID As String
    Dim strURL As String
    Dim rsTmp As New ADODB.Recordset
    Dim dblInit As Double
    Dim FrameCount As Integer
    Dim iCols As Integer, iRows As Integer
    Dim objFile As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim blnFirst As Boolean
    Dim strCachePath As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strGetFilesName As String
    
    Dim aFiles() As String
    
    Dim bln1stDev As Boolean
    bln1stDev = True
    ReDim Preserve aFiles(0) As String
    
    On Error GoTo DBError
    Screen.MousePointer = vbHourglass
    
    strFTPHost = "": strDicomPath = ""
    
    strCachePath = App.Path & "\TmpImage\"
    If Not objFileSystem.FolderExists(strCachePath) Then objFileSystem.CreateFolder strCachePath
            
    strSQL = "Select A.图像号,D.用户名 As User1,D.密码 As Pwd1," & _
        "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As FtpPath1,d.设备号 as 设备号1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/' As Path1,A.图像UID As URL1, " & _
        "E.用户名 As User2,E.密码 As Pwd2," & _
        "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As FtpPath2," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/' As Path2,A.图像UID As URL2,e.设备号 as 设备号2 " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) " & _
        "And C.检查UID=[1] "
    If Len(strSerials) = 0 Then
        strSQL = strSQL & "Order By A.图像号"
    Else
        strSQL = strSQL & "And A.序列UID In(" & strSerials & ") Order By b.序列号,A.图像号"
    End If
    If blnMoved Then
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    
    Set rsTmp = OpenSQLRecord(strSQL, "获取影像文件", CheckUID)

    If rsTmp.RecordCount > 0 Then
        ClearCacheFolder strCachePath
        MkLocalDir strCachePath & rsTmp("Path1")
        blnFirst = True
        Do While Not rsTmp.EOF
            strFTPHost = "ftp://" & rsTmp("Host1"): strDicomPath = rsTmp("FtpPath1") & rsTmp("Path1")
            strLocalPath = rsTmp("Path1"): strFTPUser = Nvl(rsTmp("User1")): strFtpPwd = Nvl(rsTmp("Pwd1"))
            
'            If Dir(strCachePath & rsTmp("Path1") & rsTmp("URL1")) = vbNullString Then
                If strDeviceNO1 <> rsTmp("设备号1") Then
                    strDeviceNO1 = rsTmp("设备号1")
                    Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
                End If

                If strDeviceNO2 <> rsTmp("设备号2") Then
                    strDeviceNO2 = rsTmp("设备号2")
                    Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
                End If
                
                '---黄捷修改于2007-1-29----------
                blnFirst = False
                If rsTmp("设备号1") <> "" Then
                    strTmpFile = strCachePath & rsTmp("Path1") & rsTmp("URL1")
                    ReDim Preserve aFiles(UBound(aFiles) + 1) As String
                    aFiles(UBound(aFiles)) = rsTmp("URL1")
                ElseIf rsTmp("设备号2") <> "" Then
                    strTmpFile = strCachePath & rsTmp("Path2") & rsTmp("URL2")
                    ReDim Preserve aFiles(UBound(aFiles) + 1) As String
                    aFiles(UBound(aFiles)) = rsTmp("URL2")
                End If
                
                '---黄捷修改于2007-1-29-----结束-----
                
                
                '--------曾超所写---黄捷修改于2007-1-29----------------
'                blnFirst = False
'                strTmpFile = strCachePath & rsTmp("Path1") & rsTmp("URL1")
'                strGetFilesName = Inet1.FuncDirFiles(strLocalPath)
'                If InStr(1, strGetFilesName, rsTmp("URL1")) > 0 Then
'                    ReDim Preserve aFiles(UBound(aFiles) + 1) As String
'                    aFiles(UBound(aFiles)) = rsTmp("URL1")
'                Else
'                    strTmpFile = strCachePath & rsTmp("Path2") & rsTmp("URL2")
'                    strGetFilesName = Inet2.FuncDirFiles(strLocalPath)
'                    ReDim Preserve aFiles(UBound(aFiles) + 1) As String
'                    aFiles(UBound(aFiles)) = rsTmp("URL2")
'                End If
                '--------曾超所写------黄捷修改于2007-1-29--------结束-----
                
'                Inet.strIPAddress = Nvl(rsTmp("Host1")): Inet.strUser = Nvl(rsTmp("User1")): Inet.strPsw = Nvl(rsTmp("Pwd1"))
'                If Inet1.FuncDownloadFile(strDicomPath, strTmpFile, rsTmp("URL1")) <> 0 Then
'                    strFtpHost = "ftp://" & Nvl(rsTmp("Host2")): strDicomPath = rsTmp("FtpPath2") & rsTmp("Path2")
'                    strLocalPath = rsTmp("Path2"): strFtpUser = Nvl(rsTmp("User2")): strFtpPwd = Nvl(rsTmp("Pwd2"))
'
'                    strTmpFile = strCachePath & rsTmp("URL2")
''                    Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
'                    If Inet2.FuncDownloadFile(strDicomPath, strTmpFile, rsTmp("URL2")) = 0 Then
'                        ReDim Preserve aFiles(UBound(aFiles) + 1) As String
'                        aFiles(UBound(aFiles)) = rsTmp("URL1")
'                    End If
'                Else
'                    ReDim Preserve aFiles(UBound(aFiles) + 1) As String
'                    aFiles(UBound(aFiles)) = rsTmp("URL1")
'                End If
'            Else
'                ReDim Preserve aFiles(UBound(aFiles) + 1) As String
'                aFiles(UBound(aFiles)) = rsTmp("URL1")
'            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    Screen.MousePointer = vbDefault
    GetAllImageFiles = aFiles
    Exit Function

ReadURLError:
    If bln1stDev Then
        bln1stDev = False
        strFTPHost = rsTmp("Host2"): strDicomPath = rsTmp("FtpPath2") & rsTmp("Path2")
        Resume
    Else
        If ErrCenter() = 1 Then Resume
        Screen.MousePointer = vbDefault
        Call SaveErrLog
    End If
    Exit Function

DBError:
    If ErrCenter() = 1 Then Resume
    Screen.MousePointer = vbDefault
    Call SaveErrLog
End Function

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'功能：当指定目录的大小达到一定百分比时，清空该目录
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

'---以下为申请登记需要
'-------------------------------------------------------------------------------------------------
Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：lngHwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：-2=未加处理,其它=匹配的索引(含不匹配的索引)
'说明：请将该函数在KeyPress事件中调用。

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '输入间隔(缺省为0.5秒)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 '使ComboBox本身的单字匹配功能失效
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '在这里对回车不作处理
    End If
End Function

Public Sub CheckLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function InputIsCard(txtInput As Object, KeyAscii As Integer) As Boolean
'功能：判断指定文本框中当前输入是否在刷卡,根据处理密文显示
    Dim strText As String, blnCard As Boolean
    Dim arrMask As Variant, i As Long

    '当前键入后显示的内容(还未显示出来)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = UCase(strText & Chr(KeyAscii))
    End If
        
    '判断是否在刷卡
    blnCard = False
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then
        blnCard = True
    ElseIf gstrCardMask <> "" Then
        arrMask = Split(gstrCardMask, "|")
        For i = 0 To UBound(arrMask)
            If strText Like arrMask(i) & "*" Then
                If IsNumeric(Mid(strText, Len(arrMask(i)) + 1)) And IsNumeric(Mid(strText, Len(arrMask(i)) + 1, 1)) Then
                    blnCard = True
                End If
            End If
        Next
    End If
    
    '刷卡时卡号是否密文显示
    If blnCard Then
        txtInput.PasswordChar = IIf(gblnCardHide, "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    
    InputIsCard = blnCard
End Function

Public Function GetSysParVal(Optional ByVal int参数号 As Integer = -9999, Optional ByVal strDefault As String) As String
'功能：获取指定系统参数的值
'参数：int参数号=为-9999时，初始化参数集
'      strDefault=如果没有值或为空的缺省值
    Dim blnDo As Boolean, strSQL As String
    
    On Error GoTo errH
    
    blnDo = True
    If Not grsSysPars Is Nothing Then
        If grsSysPars.State = 1 Then blnDo = False
    End If
    If blnDo Then
        strSQL = "Select 参数号,参数名,参数值 From 系统参数表"
        Set grsSysPars = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(grsSysPars, strSQL, "GetSysParVal")
    End If
    
    If int参数号 <> -9999 Then
        grsSysPars.Filter = "参数号=" & int参数号
        If Not grsSysPars.EOF Then
            GetSysParVal = Nvl(grsSysPars!参数值, strDefault)
        Else
            GetSysParVal = strDefault
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetTrayHeight() As Long
    '------------------------------------------------------------------------------------------------------------------
    '功能:获取任务栏的高度
    '------------------------------------------------------------------------------------------------------------------
    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    Call GetWindowRect(lngHwd, objRect)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
End Function

Private Function funAssembleImage(AssembleViewer As DicomImages, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As DicomImage

'组合viewer中的显示的所有图像成一个图像

    Dim Image As New DicomImage '新图像
    Dim imgs As New DicomImages '临时存储屏幕采集的图像集
    Dim intWidth As Integer     '新图像的宽度
    Dim intHeight As Integer    '新图像的高度
    Dim Simg As New DicomImage
    Dim intLeft As Integer
    Dim intRight As Integer
    Dim intTop As Integer
    Dim intBottom As Integer
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '单张图像可占用的区域宽度
    Dim intImgRectHeight As Integer '单张图像可占用的区域高度
    Dim i As Integer
    Dim intMaxWidth As Integer      '拼接后图像的最大宽度
    Dim intMaxHeight As Integer     '拼接后图像的最大高度
    Dim intBorder As Integer        '图像之间的边距
    Dim intImgX As Integer          'X方向的图像数量
    Dim intImgY As Integer          'Y方向的图像数量
    Dim intActualSizex As Integer   '图像旋转变换后X方向的像素点数
    Dim intActualSizey As Integer   '图像旋转变换后Y方向的像素点数
    Dim intOffsetX As Integer       '拼接时X方向的位移
    Dim intOffsetY As Integer       '拼接时Y方向的位移
    Dim dlImgLabel As DicomLabel    '图像的标注
    Dim lngWhiteX As Long           '将图象底色改成白色的X宽度
    Dim lngWhiteY As Long           '将图象底色改成白色的Y高度
    
    If AssembleViewer.Count <= 0 Then
        '返回一个黑图**************
        Exit Function
    End If

    '计算新图像的宽度和高度

    '新图像的宽度和高度不能够大于intMaxWidth×intMaxHeight（宽度×高度）
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '估算新图像的宽度和高度

    '使用原图像的宽度和高度和，并用Viewer的比例来修正。

    '估算图像的新宽高

    For i = 1 To AssembleViewer.Count
        sZoom = (lngWidth / intCols) / (AssembleViewer(i).SizeX * Screen.TwipsPerPixelX)
        If sZoom > (lngHeight / intRows) / (AssembleViewer(i).SizeY * Screen.TwipsPerPixelY) Then
            sZoom = (lngHeight / intRows) / (AssembleViewer(i).SizeY * Screen.TwipsPerPixelY)
        End If
        AssembleViewer(i).Zoom = sZoom
        '采集图像
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '精确计算新图像的宽度和高度
    intImgRectWidth = 0
    intImgRectHeight = 0

    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).SizeX Then intImgRectWidth = imgs(i).SizeX
        If intImgRectHeight < imgs(i).SizeY Then intImgRectHeight = imgs(i).SizeY
        imgs(i).Attributes.Add &H8, &H16, "doSOP_SecondaryCapture"
    Next i
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows

    '创建新图像
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT都是MONOCHROME2,CR都是MONOCHROME1？
    Image.Attributes.Add &H28, &H10, intHeight  'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth * 3, intHeight * 3) As Byte
    For lngWhiteX = 0 To intWidth * 3
        For lngWhiteY = 0 To intHeight * 3
            pix(lngWhiteX, lngWhiteY) = 255
        Next lngWhiteY
    Next lngWhiteX
    Image.Attributes.Add &H7FE0, &H10, pix

    '拼接新图像
    For i = 1 To imgs.Count
        '计算图像内位移
        intOffsetX = (intImgRectWidth - imgs(i).SizeX - intBorder) / 2
        intOffsetY = (intImgRectHeight - imgs(i).SizeY - intBorder) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod intCols) * intImgRectWidth + intOffsetX, ((i - 1) \ intCols) * intImgRectHeight + intOffsetY, imgs(i).SizeX, imgs(i).SizeY, 1, 1, 1, False
    Next i

    Set funAssembleImage = Image
End Function

Private Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
    Dim iCols As Integer, iRows As Integer
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1

    Do While iRows * iCols > ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols - 1
        Else
            iRows = iRows - 1
        End If
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MaxRows > 0 And iRows > MaxRows Then
        iRows = MaxRows
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MaxRows > 0 And iRows > MaxRows Then iRows = MaxRows
    
    Rows = iRows: Cols = iCols
End Sub

