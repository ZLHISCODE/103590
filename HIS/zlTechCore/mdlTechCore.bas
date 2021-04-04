Attribute VB_Name = "mdlTechCore"
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
Public gbln加班加价 As Boolean
Public grsSysPars As ADODB.Recordset
Public grsDuty As ADODB.Recordset '存放医生职务
Public gstr动态费别 As String '存放门诊当前科室可用动态费别,在公共函数中使用,使用时才赋值:CalcDrugPrice,CalcPrice

'医保变量
Public gclsInsure As New clsInsure

'CIS系统参数
Public gbln药品按规格下医嘱 As Boolean
Public gint过敏登记有效天数 As Integer
Public gbln长期医嘱次日生效 As Boolean
Public gstr发送划价单 As String
Public gbln药疗划价单 As Boolean
Public gbln其他划价单 As Boolean
Public gbln执行后审核 As Boolean
Public gintRXCount As Integer

'HIS系统参数
Public gbln中医 As Boolean '是否使用中医病案
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"

Public gint挂号天数 As Integer '挂号单有效天数
Public gbln收费类别 As Boolean '是否首先输入类别
Public gbln商品名 As Boolean '西成药是否按商品名显示
Public gbln住院自动发料 As Boolean '住院记帐完成后是否自动发料
Public gbln门诊自动发料 As Boolean '门诊记帐完成后是否自动发料
Public gbln从项汇总折扣 As Boolean '从属项目汇总计算折扣
Public gbln报警包含划价费用 As Boolean '记帐报警包含划价费用
Public gbln病区科室独立 As Boolean '病区和科室是否独立管理
Public gint诊疗编码 As Integer '0-顺序编号,1-种类+分类号+顺序编号
Public gint诊断来源 As Integer '1-由医生选择输入来源,2-按照诊断标准输入,3-按照疾病编码输入
Public gint诊断输入 As Integer '1-允许自由输入,2-从数据库提取输入,3-仅医保病人从数据库输入
Public gstrMatchMode As String '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
Public gbyt检查未执行 As Byte '出院转科时是否检查有未执行项目及未发药品:0-不检查,1-检查并提示,2-检查并禁止
Public gint医保对码 As Integer '是否对住院医保病人的项目对码情况进行检查:0-不检查,1-检查并提醒,2-检查并禁止

'医技工作站系统费用参数
Public gbytBillOpt As Byte '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
Public gblnStock As Boolean '指定药房时是否限定输入药品的库存
Public gstr医保费用类型 As String '医保病人允许的费用类型
Public gstr公费费用类型 As String '公费病人允许的费用类型

'电子签名
Public gintCA As Integer '电子签名认证中心
Public gstrESign As String '电子签名控制场合
Public gobjESign As Object '电子签名接口部件

'-----------------------------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    部门ID As Long
    部门码 As String
    部门名 As String
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
    support负数记帐 = 35            '是否允许负数记帐，操作员首先要拥有负数记帐的权限。此参数缺省为真，不支持的接口需单独处理
End Enum
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

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
    Call SQLTest(App.ProductName, "mdlTechCore", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIF(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIF(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIF(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        UserInfo.用户名 = IIF(IsNull(rsTmp!用户名), "", rsTmp!用户名)
        gstrDBUser = UserInfo.用户名
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
        strSQL = "Select Max(Length(床号)) as 长度 From 床位状况记录 Where 状态='占用' And 病区ID" & IIF(lng部门ID = 0, " is Not NULL", "=" & lng部门ID)
    Else
        strSQL = "Select Max(Length(床号)) as 长度 From 床位状况记录 Where 状态='占用' And 科室ID" & IIF(lng部门ID = 0, " is Not NULL", "=" & lng部门ID)
    End If
    
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, "mdlCISWork", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    If Not rsTmp.EOF Then GetMaxBedLen = IIF(IsNull(rsTmp!长度), 0, rsTmp!长度)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病区ID(Optional ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, _
    Optional ByVal lng科室ID As Long) As Long
'功能：根据科室ID或病人获取对应的病区ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng病人ID <> 0 And lng主页ID <> 0 Then
        strSQL = "Select 当前病区ID as 病区ID From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng病人ID, lng主页ID)
    Else
        strSQL = "Select 病区ID From 床位状况记录 Where 科室ID=[1] Group by 病区ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng科室ID)
    End If
    If Not rsTmp.EOF Then Get病区ID = Nvl(rsTmp!病区ID, 0)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
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

Public Function InitStockCheck(ByVal int范围 As Integer, Optional ByVal bln卫材 As Boolean) As Collection
'功能：读取不同库房出库检查方式于集合中
'参数：int范围=1-门诊,2-住院
'      bln卫材=是否包含卫材的发料部门
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
        " And B.工作性质 in('中药房','西药房','成药房'" & IIF(bln卫材, ",'发料部门'", "") & ")" & _
        " And C.库房ID(+)=A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int范围)
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

Public Function GetSysParVal(Optional ByVal int参数号 As Integer = -9999, Optional ByVal strDefault As String) As String
'功能：获取指定系统参数的值
'参数：int参数号=为-9999时，初始化参数集
'      strDefault=如果没有值或为空的缺省值
    Dim blnDo As Boolean, strSQL As String
    
    On Error GoTo errH
    
    blnDo = True
    If int参数号 <> -9999 Then
        If Not grsSysPars Is Nothing Then
            If grsSysPars.State = 1 Then blnDo = False
        End If
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

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    'HIS系统参数
    '---------------------------------------------------------
    Call GetSysParVal
    
    '费用金额小数点位数
    gbytDec = 2: gstrDec = "0.00"
    grsSysPars.Filter = "参数号=9"
    If Not grsSysPars.EOF Then
        gbytDec = Val(Nvl(grsSysPars!参数值, 2))
        gstrDec = "0." & String(gbytDec, "0")
    End If
    
    '指定药房时限制库存
    grsSysPars.Filter = "参数号=18"
    If Not grsSysPars.EOF Then gblnStock = Nvl(grsSysPars!参数值, 0) <> 0
    
    '挂号有效天数
    grsSysPars.Filter = "参数号=21"
    If Not grsSysPars.EOF Then gint挂号天数 = Nvl(grsSysPars!参数值, 0)
    
    '检查未执行项目
    grsSysPars.Filter = "参数号=22"
    If Not grsSysPars.EOF Then gbyt检查未执行 = Nvl(grsSysPars!参数值, 0)
    
    '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
    grsSysPars.Filter = "参数号=23"
    If Not grsSysPars.EOF Then gbytBillOpt = Nvl(grsSysPars!参数值, 0)
    
    '电子签名认证中心
    grsSysPars.Filter = "参数号=25"
    If Not grsSysPars.EOF Then gintCA = Val(Nvl(grsSysPars!参数值, "0"))
    
    '电子签名控制场合
    grsSysPars.Filter = "参数号=26"
    If Not grsSysPars.EOF Then gstrESign = Nvl(grsSysPars!参数值)
    
    '是否使用中医
    grsSysPars.Filter = "参数号=31"
    If Not grsSysPars.EOF Then gbln中医 = Nvl(grsSysPars!参数值, 0) <> 0
    
    '医保费用类型
    grsSysPars.Filter = "参数号=41"
    If Not grsSysPars.EOF Then
        gstr医保费用类型 = "'" & Replace(Nvl(grsSysPars!参数值), "|", "','") & "'"
    End If

    '公费费用类型
    grsSysPars.Filter = "参数号=42"
    If Not grsSysPars.EOF Then
        gstr公费费用类型 = "'" & Replace(Nvl(grsSysPars!参数值), "|", "','") & "'"
    End If
    
    '门诊处方条数限制
    grsSysPars.Filter = "参数号=56"
    If Not grsSysPars.EOF Then gintRXCount = Val(Nvl(grsSysPars!参数值, 0))

    '医保对码检查
    grsSysPars.Filter = "参数号=59"
    gint医保对码 = 1
    If Not grsSysPars.EOF Then gint医保对码 = Val(Nvl(grsSysPars!参数值, 1))

    '诊疗编码递增模式
    grsSysPars.Filter = "参数号=61"
    If Not grsSysPars.EOF Then gint诊疗编码 = Val(Nvl(grsSysPars!参数值, 0))
    
    '住院自动发料
    grsSysPars.Filter = "参数号=63"
    If Not grsSysPars.EOF Then
        gbln住院自动发料 = Nvl(grsSysPars!参数值, 0) <> 0
    End If
    
    '药品按规格下医嘱
    grsSysPars.Filter = "参数号=69"
    If Not grsSysPars.EOF Then gbln药品按规格下医嘱 = Val(Nvl(grsSysPars!参数值, 0)) = 1
    
    '皮试结果有效时间
    grsSysPars.Filter = "参数号=70"
    If Not grsSysPars.EOF Then gint过敏登记有效天数 = Val(Nvl(grsSysPars!参数值, 0))
    
    '长期医嘱次日生效
    grsSysPars.Filter = "参数号=71"
    If Not grsSysPars.EOF Then gbln长期医嘱次日生效 = Val(Nvl(grsSysPars!参数值, 0)) = 1
    
    '是否要求首先输入类别
    grsSysPars.Filter = "参数号=72"
    If Not grsSysPars.EOF Then gbln收费类别 = Nvl(grsSysPars!参数值, 1) <> 0
    
    '西成药是否按商品名显示
    grsSysPars.Filter = "参数号=74"
    If Not grsSysPars.EOF Then gbln商品名 = Nvl(grsSysPars!参数值, 0) <> 0
    
    '药疗生成划价单
    grsSysPars.Filter = "参数号=79"
    If Not grsSysPars.EOF Then gbln药疗划价单 = Nvl(grsSysPars!参数值, 0) <> 0
    
    '其他生成划价单
    grsSysPars.Filter = "参数号=80"
    If Not grsSysPars.EOF Then gbln其他划价单 = Nvl(grsSysPars!参数值, 0) <> 0
    
    '执行后自动审核
    grsSysPars.Filter = "参数号=81"
    If Not grsSysPars.EOF Then gbln执行后审核 = Nvl(grsSysPars!参数值, 0) <> 0
            
    '门诊自动发料
    grsSysPars.Filter = "参数号=92"
    If Not grsSysPars.EOF Then
        gbln门诊自动发料 = Nvl(grsSysPars!参数值, 0) <> 0
    End If
    
    '从属项目汇总计算折扣
    grsSysPars.Filter = "参数号=93"
    If Not grsSysPars.EOF Then gbln从项汇总折扣 = Nvl(grsSysPars!参数值, 0) <> 0
    
    '记帐报警包含划价费用
    grsSysPars.Filter = "参数号=98"
    If Not grsSysPars.EOF Then gbln报警包含划价费用 = Nvl(grsSysPars!参数值, 0) <> 0
    
    '病区和科室是否独立管理
    grsSysPars.Filter = "参数号=99"
    If Not grsSysPars.EOF Then gbln病区科室独立 = Nvl(grsSysPars!参数值, 0) <> 0
    
    '电子签名初始化:只要使用即出来
    If gintCA <> 0 Then
        If gobjESign Is Nothing Then
            On Error Resume Next
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            Err.Clear: On Error GoTo 0
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle)
            End If
        End If
    Else
        Set gobjESign = Nothing
    End If
    
    InitSysPar = True
End Function

Public Function GetPatiYear(lng病人ID As Long) As Integer
'功能：获取病人的准确年龄
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intYear As Integer
    
    On Error GoTo errH
    
    strSQL = "Select Sysdate as 当前,出生日期,年龄 From 病人信息 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng病人ID)
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

Public Function GET部门名称(lngID As Long, Optional ByRef rs部门 As ADODB.Recordset) As String
'功能：获取部门名称
'参数：lngID=部门ID
'返回：部门名称
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If rs部门 Is Nothing Then
        strSQL = "Select 名称 from 部门表 Where ID=" & lngID
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlCISWork")
    Else
        Set rsTmp = rs部门
        rsTmp.Filter = "ID=" & lngID
        If rsTmp.RecordCount = 0 Then
            strSQL = "Select 名称 from 部门表 Where ID=" & lngID
            Set rsTmp = New ADODB.Recordset
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlCISWork")
        End If
    End If
    
    If Not rsTmp.EOF Then GET部门名称 = rsTmp!名称
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", CStr(int类型), int来源)
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
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlCISWork")
    If Not rsTmp.EOF Then
        Check上班安排 = Nvl(rsTmp!Num, 0) > 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get收费执行科室ID(ByVal lng病人ID As Long, lng主页ID As Long, _
    ByVal str类别 As String, ByVal lng项目ID As Long, ByVal int执行科室 As Integer, _
    ByVal lng病人科室ID As Long, ByVal lng开单科室ID As Long, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal lng发料部门 As Long) As Long
'功能：获取非药收费项目的执行科室
'参数：int范围=1.门诊,2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim str药房 As String, lng药房 As Long
    Dim bytDay As Byte, strIDs As String
    
    On Error GoTo errH
    
    If str类别 = "4" Then
        '列出所有发料部门时
'        strSQL = "Select B.服务对象,A.编码,A.ID From 部门表 A,部门性质说明 B" & _
'            " Where A.ID=B.部门ID And B.工作性质='发料部门' And B.服务对象 IN([1],3)" & _
'            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
'            " Order by B.服务对象,A.编码"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int范围)
'        If Not rsTmp.EOF Then
'            If lng发料部门 <> 0 Then rsTmp.Filter = "ID=" & lng发料部门
'            If rsTmp.EOF Then rsTmp.Filter = 0
'            Get收费执行科室ID = rsTmp!ID
'        End If
    
        '有执行科室设置时
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int范围, lng病人科室ID, lng项目ID)
        If Not rsTmp.EOF Then
            rsTmp.Filter = "开单科室ID=" & lng病人科室ID
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
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(int范围 = 2, "住院", "门诊") & "缺省西药房", 0))
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(int范围 = 2, "住院", "门诊") & "缺省成药房", 0))
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(int范围 = 2, "住院", "门诊") & "缺省中药房", 0))
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str药房, int范围, lng病人科室ID, lng项目ID, bytDay)
        If Not rsTmp.EOF Then
            strIDs = ""
            rsTmp.Filter = "开单科室ID=" & lng病人科室ID
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
                Get收费执行科室ID = lng病人科室ID
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    Get收费执行科室ID = lng病人科室ID
                Else
                    Get收费执行科室ID = Get病区ID(lng病人ID, lng主页ID)
                End If
            Case 3 '3-操作员所在科室
                Get收费执行科室ID = UserInfo.部门ID
            Case 4 '4-指定科室
                strSQL = "Select Nvl(开单科室ID,0) as 开单科室ID,执行科室ID" & _
                    " From 收费执行科室 Where 收费细目ID=[1]" & _
                    " And (病人来源 is NULL Or 病人来源=[2])" & _
                    " And (开单科室ID is NULL Or 开单科室ID=[3])" & _
                    " Order by Decode(病人来源,Null,2,1)" '默认科室优先
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID, int范围, lng病人科室ID)
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "开单科室ID=" & lng病人科室ID
                    If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0"
                    If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
                End If
            Case 6 '6-开单人所在科室
                Get收费执行科室ID = lng开单科室ID
        End Select
        If Get收费执行科室ID = 0 Then Get收费执行科室ID = UserInfo.部门ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function




Public Function Get诊疗执行科室ID(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str类别 As String, ByVal lng项目ID As Long, ByVal lng药品ID As Long, ByVal int执行科室 As Integer, _
    ByVal lng病人科室ID As Long, ByVal lng开嘱科室ID As Long, ByVal int期效 As Integer, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal blnBy缺省 As Boolean) As Long
'功能：根据诊疗项目执行科室信息返回缺省的执行科室ID
'参数：lng药品ID=药品ID,确定到规格时要用
'      int执行科室=项目执行科室标志
'      lng病人科室ID=病人科室ID
'      lng西药房,lng成药房,lng中药房=药品缺省药房,药品类时需要
'      int期效=0-长嘱,1-临嘱,临嘱可能要判断上班时间
'      int范围=1-门诊,2-住院(缺省)
'      blnBy缺省=获取缺省药房时，如果本地有指定，是否按本地缺省指定的药房来，没有则不返回
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
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(int范围 = 2, "住院", "门诊") & "缺省西药房", 0))
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(int范围 = 2, "住院", "门诊") & "缺省成药房", 0))
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            lng药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(int范围 = 2, "住院", "门诊") & "缺省中药房", 0))
        End If
        
        '药品从系统指定的储备药房中找
        If int范围 = 1 Then bln上班安排 = Check上班安排(True) '住院医嘱不管药房上班安排
        If Not bln上班安排 Then
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                 IIF(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[6]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                IIF(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & _
                " Order by B.服务对象,C.编码"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str药房, int范围, lng病人科室ID, lng药品ID, lng项目ID, bytDay)
        If Not rsTmp.EOF Then
            If blnBy缺省 And lng药房 <> 0 Then
                rsTmp.Filter = "执行科室ID=" & lng药房
            Else
                Get诊疗执行科室ID = rsTmp!执行科室ID
                rsTmp.Filter = "执行科室ID=" & lng药房
                If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            End If
            If Not rsTmp.EOF Then Get诊疗执行科室ID = rsTmp!执行科室ID
        End If
    Else
        Select Case int执行科室
            Case 0 '0-无执行的叮嘱
                Exit Function
            Case 1 '1-病人所在科室
                Get诊疗执行科室ID = lng病人科室ID
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    Get诊疗执行科室ID = lng病人科室ID
                Else
                    Get诊疗执行科室ID = Get病区ID(lng病人ID, lng主页ID)
                End If
            Case 3 '3-操作员所在科室
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
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID, int范围, lng病人科室ID, bytDay)
                If Not rsTmp.EOF Then
                    Get诊疗执行科室ID = rsTmp!执行科室ID
                    rsTmp.Filter = "开单科室ID=" & lng病人科室ID
                    If Not rsTmp.EOF Then Get诊疗执行科室ID = rsTmp!执行科室ID
                End If
            Case 5 '5-院外执行
                Exit Function
            Case 6 '6-开单人所在科室
                Get诊疗执行科室ID = lng开嘱科室ID
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Have部门性质(ByVal lng科室ID As Long, ByVal str性质 As String) As Boolean
'功能：检查指定科室是否具有指定工作性质
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select 部门ID From 部门性质说明 Where 部门ID=[1] And 工作性质=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng科室ID, str性质)
    Have部门性质 = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Load动态费别(lng科室ID As Long) As String
'功能：权限指定科室读取当前有效的动态费别(目前只用于门诊)
'返回：费别串="三八节,五一节"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select 编码,简码,名称 From 费别" & _
        " Where Nvl(属性,1)=2 And Nvl(适用科室,1)=1 And Nvl(服务对象,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(有效开始,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(有效结束,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Union ALL" & _
        " Select Distinct A.编码,A.简码,A.名称" & _
        " From 费别 A,费别适用科室 B" & _
        " Where A.名称=B.费别 And B.科室ID=[1]" & _
        " And Nvl(A.属性,1)=2 And Nvl(A.适用科室,1)=2 And Nvl(A.服务对象,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(A.有效开始,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(A.有效结束,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Load动态费别", lng科室ID)
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    Load动态费别 = Mid(strTmp, 2)
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
            " From " & IIF(lng药品ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
            " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
            IIF(lng药品ID <> 0, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]")
    Else
        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
        strSQL = _
            " Select Distinct C.ID" & _
            " From " & IIF(lng药品ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C,部门安排 D" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
            " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And D.部门ID=C.ID And D.星期=[6]" & _
            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
            " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
            IIF(lng药品ID <> 0, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]")
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str药房, int范围, lng科室ID, lng药品ID, lng项目ID, bytDay)
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

Public Function Get诊疗执行科室(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    objCbo As Object, ByVal str类别 As String, ByVal lng项目ID As Long, ByVal lng药品ID As Long, _
    ByVal int执行科室 As Integer, ByVal lng病人科室ID As Long, ByVal lng开嘱科室ID As Long, _
    ByVal lng当前执行ID As Long, ByVal int期效 As Integer, Optional ByVal int范围 As Integer = 2) As Boolean
'功能：根据诊疗项目执行科室信息返回可用的执行科室在指定下拉框中
'参数：int执行科室=项目执行科室标志
'      lng病人科室ID=病人科室ID
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
                " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                IIF(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象" & _
                " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[8]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                IIF(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & _
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
                        " From 部门表 A,病案主页 B" & _
                        " Where A.ID=B.当前病区ID And B.病人ID=[9] And B.主页ID=[10]" & _
                        " Union " & _
                        " Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                        " Order by 编码"
                End If
            Case 3 '3-操作员所在科室
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
            Case 6 '6-开单人所在科室
                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([11],[6]) Order by 编码"
        End Select
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str药房, int范围, lng病人科室ID, lng药品ID, lng项目ID, lng当前执行ID, UserInfo.部门ID, bytDay, lng病人ID, lng主页ID, lng开嘱科室ID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str编码)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str频率, "," & str范围 & ",")
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int范围)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int范围)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", int范围, lng给药途径ID, str频率)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str频率, lng给药途径ID, int范围)
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
    Optional objCbo As Object, Optional ByVal int范围 As Integer = 2, Optional blnAppend As Boolean) As Boolean
'功能：获取可用的开嘱医生在指定的下拉框中
'参数：lng病人科室ID=病人所在科室ID
'      bln护士站=是否由护士代医生下医嘱
'      objCbo=要加入医生清单的下拉框
'      str缺省医生=缺省定位的医生,如果不传objCbo,则先优先定位,再返回缺省医生和医生ID
'      int范围=1-门诊,2-住院(缺省)
'      blnAppend=是否将当前缺省医生附加到列表中的形式,此时"bln护士站,str缺省医生,objCbo"都有
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    If bln护士站 Then
        If blnAppend And str缺省医生 <> "" Then
            strSQL = "Select ID,编号,姓名,简码 From 人员表 Where 姓名=[4]"
        Else
            '病人所在科室的医生
            strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIF(objCbo Is Nothing, ",B.部门ID", "") & _
                " From 人员表 A,部门人员 B,人员性质说明 C" & _
                " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
                " And B.部门ID=[1]" & _
                " Order by A.简码"
            '全院住院科室的医生
            strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN([2],3)"
            strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIF(objCbo Is Nothing, ",B.部门ID", "") & _
                " From 人员表 A,部门人员 B,人员性质说明 C" & _
                " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
                " And B.部门ID IN(" & strSQL & ")" & _
                " Order by A.简码"
        End If
    Else '医生下医嘱时,限制为只能为医生本人
        strSQL = "Select ID,编号,姓名,简码 From 人员表 Where ID=[3]"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng病人科室ID, int范围, UserInfo.ID, str缺省医生)
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
        If blnAppend Then
            '先删除"其它"
            i = SeekCboIndex(objCbo, -1)
            If i <> -1 Then objCbo.RemoveItem objCbo.ListCount - 1
            
            '定位或加入其它选项
            If Not rsTmp.EOF Then
                i = SeekCboIndex(objCbo, rsTmp!ID)
                If i = -1 Then
                    objCbo.AddItem Nvl(rsTmp!简码) & "-" & rsTmp!姓名
                    objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
                    Call zlControl.CboSetIndex(objCbo.Hwnd, objCbo.NewIndex)
                Else
                    Call zlControl.CboSetIndex(objCbo.Hwnd, i)
                End If
            End If
            
            '加入"其它"供选择
            AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, "[其它...]"
            SetComboData objCbo.Hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
        Else
            '全部新加入
            objCbo.Clear
            For i = 1 To rsTmp.RecordCount
                AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, Nvl(rsTmp!简码) & "-" & rsTmp!姓名
                SetComboData objCbo.Hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
                If rsTmp!姓名 = str缺省医生 Then
                    Call zlControl.CboSetIndex(objCbo.Hwnd, i - 1)
                End If
                rsTmp.MoveNext
            Next
        End If
    End If
    Get开嘱医生 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医生ID)
    
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID, lng执行科室ID)
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
    
    strTmp = IIF(int范围 = 1, "门诊", "住院")
    
    '获取药品库存(不分批或分批药品),药房不分批药品不管效期
    strSQL = _
        " Select Nvl(Sum(A.可用数量),0)/Nvl(B." & strTmp & "包装,1) as 库存" & _
        " From 药品库存 A,药品规格 B" & _
        " Where A.药品ID=B.药品ID(+) And A.性质=1" & _
        " And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
        " And A.药品ID=[1] And A.库房ID=[2]" & _
        " Group by Nvl(B." & strTmp & "包装,1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng药品ID, lng库房ID)
    If Not rsTmp.EOF Then
        GetStock = Format(rsTmp!库存, "0.00000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetItemField(ByVal strTable As String, ByVal lngID As Long, Optional ByVal strField As String) As Variant
'功能：获取指定表指定字段信息
'说明：未处理NULL值
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If strField = "" Then
        strSQL = "Select * From " & strTable & " Where ID=[1]"
    Else
        strSQL = "Select " & strField & " From " & strTable & " Where ID=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngID)
    If Not rsTmp.EOF Then
        If strField = "" Then
            Set GetItemField = rsTmp
        Else
            GetItemField = rsTmp.Fields(strField).Value
        End If
    End If
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
        IIF(bln期效 And int来源 = 1, " And A.期效=1", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng组合ID, int来源)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng配方ID, int来源)
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1)
    Else
        strSQL = _
            " Select * From 病历文件目录 Where 种类=[1] And " & _
            IIF(lngDeptID = -1, "应用=1", "应用=2 And ','||科室ID||',' Like [2]") & _
            " Order by 编号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1, "%," & lngDeptID & ",%")
        If rsTmp.EOF Then  '指定科室无该类病历，则查公用病历
            strSQL = "Select * From 病历文件目录 Where 种类=[1] And 应用=1 Order by 编号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", iFileType + 1)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngPatientID, vPageID)
    
    Do While Not rsTmp.EOF
        zlDatabase.ExecuteProcedure "ZL_病人病历_归档(" & rsTmp(0) & ",'" + UserInfo.姓名 + "')", ""
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng药房ID, lng药品ID)
    
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
                    IIF(gbln加班加价, "*Decode(A.加班加价,1,1+Nvl(B.加班加价率,0)/100,1)", "") & " as 金额" & _
                " From 收费项目目录 A,收费价目 B" & _
                " Where A.ID=B.收费细目ID And A.ID=[1]" & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng药品ID, dbl时价)
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
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", Val(Format(dbl时价 * dbl数量, gstrDec)), str费别, Abs(dbl时价), lng药品ID)
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
    Optional ByVal dbl数量 As Double, Optional ByVal blnNone加班加价 As Boolean, Optional ByVal lng执行科室ID As Long) As Double
'功能：获取收费细目的当前售价价格金额,变价返回0
'参数：str费别=是否按费别计算打折的实收金额
'      dbl数量=按费别计算时,必须要传入数量(按售价单位),这时计算的是实收金额
'      lng执行科室ID=当传入了费别时需要,可能按成本加打折计算
'      gbln加班加价=发送时计算才有用,其它地方都为False
'      blnNone加班加价=为真时,"gbln加班加价"无效
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dbl金额 As Double
    
    On Error GoTo errH
    
    If str费别 = "" Then
        strSQL = _
            " Select Sum(Decode(Nvl(A.是否变价,0),1,NULL," & _
                "B.现价" & IIF(gbln加班加价 And Not blnNone加班加价, "*Decode(A.加班加价,1,1+Nvl(B.加班加价率,0)/100,1)", "") & ")) as 金额" & _
            " From 收费项目目录 A,收费价目 B" & _
            " Where A.ID=B.收费细目ID And A.ID=[1]" & _
            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
        If Not rsTmp.EOF Then dbl金额 = Nvl(rsTmp!金额, 0)
    Else
        '本来可以将ActualMoney函数的SQL一起写在这里，但费别可能被删除而求不出数据
        strSQL = _
            " Select A.屏蔽费别,A.加班加价,B.加班加价率,B.收入项目ID,Decode(Nvl(A.是否变价,0),1,NULL," & _
                "B.现价" & IIF(gbln加班加价 And Not blnNone加班加价, "*Decode(A.加班加价,1,1+Nvl(B.加班加价率,0)/100,1)", "") & ") as 金额" & _
            " From 收费项目目录 A,收费价目 B" & _
            " Where A.ID=B.收费细目ID And A.ID=[1]" & _
            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
        For i = 1 To rsTmp.RecordCount
            If Nvl(rsTmp!屏蔽费别, 0) = 1 Then
                dbl金额 = dbl金额 + Format(dbl数量 * Format(Nvl(rsTmp!金额, 0), "0.00000"), gstrDec)
            Else
                dbl金额 = dbl金额 + ActualMoney(str费别 & IIF(gstr动态费别 <> "", "," & gstr动态费别, ""), rsTmp!收入项目ID, Format(dbl数量 * Format(Nvl(rsTmp!金额, 0), "0.00000"), gstrDec), _
                    lng项目ID, lng执行科室ID, dbl数量, IIF(gbln加班加价 And Not blnNone加班加价 And Nvl(rsTmp!加班加价, 0) = 1, Nvl(rsTmp!加班加价率, 0) / 100, 0))
            End If
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


Public Function ActualMoney(str费别 As String, ByVal lng收入项目ID As Long, ByVal cur应收金额 As Currency, _
    Optional ByVal lng药品ID As Long, Optional ByVal lng库房ID As Long, Optional ByVal dbl数量 As Double, Optional ByVal dbl加班加价率 As Double) As Currency
'功能：根据费别,收入项目ID,应收金额,按分段比例打折规则计算实收金额；或根据药品相关信息按成本加收比例规则计算实收金额
'参数：str费别=病人费别；如果是按动态费别计算,传入格式为"病人费别,动态费别1,动态费别2,..."
'      str类别,lng药品ID,lng库房ID,dbl数量,dbl加班加价率=药品类项目需要传入
'      dbl数量=包含付数在内的售价数量
'      dbl加班加价率=小数比率,传入的应收金额已按加班加价计算时需要，用于还原及重算
'返回：按打折规则和比例计算的实收金额,如果是动态费别,则"str费别"返回最优惠费别(注意如果未打折计算,可能原样返回)
'说明：
'按成本价加收比例打折的两种计算方法(实际是一种)：
'1.打折金额 = 成本金额 * (1 + 加收比例)
'2.打折金额 = 成本价 * (1 + 加收比例) * 零售数量
'相关的计算公式：
'      成本价 = 药品售价 * (1 - 差价率)
'      成本金额 = 售价金额 * (1 - 差价率) = 成本价 * 零售数量
'      有库存金额时:差价率 = 库存差价 / 库存金额,否则:差价率 = 指导差价率
'      对于分批药品，应每个出库批次分别计算成本价和成本金额
'        对于时价分批，"药品售价=实际金额/实际数量"；分批或时价药品库存不足时，不予打折计算。
    Dim rsTmp As New ADODB.Recordset
    Dim rsBase As New ADODB.Recordset
    Dim rsDrug As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dblCost As Double, dblRate As Double
    Dim dblCurTime As Double, dblRebate As Double
    
    Dim blnDynamic As Boolean, dblCurMoney As Double
    Dim strMin费别 As String, dblMinMoney As Double
    
    On Error GoTo errH
    
    '设置不打折计算时的缺省值
    ActualMoney = cur应收金额
    If str费别 = "" Or cur应收金额 = 0 Then Exit Function
       
    blnDynamic = InStr(str费别, ",") > 0
    If Not blnDynamic Then
        strSQL = _
            " Select 费别,Nvl(实收比率,0) as 实收比率,[3]*Nvl(实收比率,0)/100 as 实收金额,Nvl(计算方法,0) as 计算方法" & _
            " From 费别明细 Where 收入项目ID=[1] and 费别=[2] and Abs([3]) Between 应收段首值 and 应收段尾值"
        Set rsBase = zlDatabase.OpenSQLRecord(strSQL, "ActualMoney", lng收入项目ID, str费别, cur应收金额)
    Else
        strSQL = _
            " Select A.费别,Nvl(A.实收比率,0) as 实收比率,[3]*Nvl(A.实收比率,0)/100 as 实收金额,Nvl(A.计算方法,0) as 计算方法" & _
            " From 费别明细 A,费别 B" & _
            " Where A.费别=B.名称 And A.收入项目ID=[1] And Instr([2],','||B.名称||',')>0" & _
            " And Abs([3]) Between A.应收段首值 and A.应收段尾值" & _
            " Order by Nvl(A.计算方法,0),Nvl(A.实收比率,0),Nvl(B.属性,1),B.编码"
        Set rsBase = zlDatabase.OpenSQLRecord(strSQL, "ActualMoney", lng收入项目ID, "," & str费别 & ",", cur应收金额)
    End If
    If rsBase.EOF Then Exit Function
    
    dblMinMoney = 9999999999999#
        
    '按分段比例打折规则计算:按比例打折规则中打折比例最小的费别
    If blnDynamic Then rsBase.Filter = "计算方法=0"
    If Not rsBase.EOF Then
        If Nvl(rsBase!计算方法, 0) = 0 Then
            strMin费别 = rsBase!费别
            dblMinMoney = rsBase!实收金额
        End If
    End If
    
    '按成本加收比例规则计算:按成本加收规则中加收比例最小的费别
    If blnDynamic Then rsBase.Filter = "计算方法=1"
    If Not rsBase.EOF Then
        If Nvl(rsBase!计算方法, 0) = 1 And lng药品ID <> 0 And dbl数量 <> 0 Then
            dblRate = Nvl(rsBase!实收比率, 0) / 100
            cur应收金额 = cur应收金额 / (1 + dbl加班加价率) '传入的应收是根据加班加价计算过的,所以需要还原
            
            '取药品信息
            strSQL = _
                " Select B.类别,A.指导差价率,Nvl(C.现价,0) as 售价," & _
                " Nvl(A.药房分批,0) as 分批,Nvl(B.是否变价,0) as 变价" & _
                " From 药品规格 A,收费项目目录 B,收费价目 C" & _
                " Where A.药品ID=B.ID And B.ID=C.收费细目ID And A.药品ID=[1]" & _
                " And Sysdate Between C.执行日期 And Nvl(C.终止日期,To_Date('3000-01-01', 'YYYY-MM-DD'))"
            Set rsDrug = zlDatabase.OpenSQLRecord(strSQL, "ActualMoney", lng药品ID)
            If rsDrug.EOF Then GoTo EndCalc
            If InStr(",5,6,7,", rsDrug!类别) = 0 Then GoTo EndCalc '非药品不打折
            
            If lng库房ID = 0 Then
                '没有确定药房时(或分离模式),分批或不分批药品都按指导差价率算
                dblCurMoney = cur应收金额 * (1 - Nvl(rsDrug!指导差价率, 0) / 100) * (1 + dblRate) * (1 + dbl加班加价率)
                If dblCurMoney < dblMinMoney Then strMin费别 = rsBase!费别: dblMinMoney = dblCurMoney
            ElseIf rsDrug!分批 = 0 Then
                '不分批药品:
                strSQL = "[1]*(1-Nvl(Decode(Sign(Nvl(A.实际金额,0)),1,A.实际差价/A.实际金额,B.指导差价率/100),0))"
                strSQL = "Select " & strSQL & " as 成本金额 From 药品库存 A,药品规格 B" & _
                         " Where A.药品ID(+)=B.药品ID And A.库房ID(+)=[2] And A.药品ID(+)=[3] And A.性质(+)=1 And B.药品ID=[3]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ActualMoney", cur应收金额, lng库房ID, lng药品ID)
                dblCost = Nvl(rsTmp!成本金额, 0)
                If dblCost <> 0 Then
                    dblCurMoney = dblCost * (1 + dblRate) * (1 + dbl加班加价率)
                    If dblCurMoney < dblMinMoney Then strMin费别 = rsBase!费别: dblMinMoney = dblCurMoney
                End If
            ElseIf rsDrug!分批 = 1 Then
                '分批药品:每个批次求成本加收
                strSQL = _
                    " Select Nvl(批次,0) as 批次,Nvl(可用数量,0) as 库存," & _
                    " Nvl(Decode(Nvl(实际数量,0),0,0,实际金额/实际数量),0) as 时价," & _
                    " Nvl(实际差价,0) as 实际差价,Nvl(实际金额,0) as 实际金额" & _
                    " From 药品库存" & _
                    " Where (Nvl(批次,0)=0 Or 效期 is NULL Or 效期>Trunc(Sysdate))" & _
                    " And Nvl(可用数量,0)<>0 And 性质=1 And 库房ID=[1] And 药品ID=[2]" & _
                    " Order by Nvl(批次,0)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ActualMoney", lng库房ID, lng药品ID)
                
                dblRebate = 0: dblCurTime = 0
                For i = 1 To rsTmp.RecordCount
                    If dbl数量 = 0 Then Exit For
                    
                    '取小者
                    If dbl数量 <= rsTmp!库存 Then
                        dblCurTime = dbl数量
                    Else
                        dblCurTime = rsTmp!库存
                    End If
                    
                    If rsTmp!实际金额 <> 0 Then
                        '本批次成本价:按库存差价率算
                        dblCost = IIF(rsDrug!变价 = 1, rsTmp!时价, rsDrug!售价) * (1 - rsTmp!实际差价 / rsTmp!实际金额)
                    Else
                        '无库存金额按指导差价率算:无库存的批次在SQL中已排开
                        dblCost = IIF(rsDrug!变价 = 1, rsTmp!时价, rsDrug!售价) * (1 - Nvl(rsDrug!指导差价率, 0) / 100)
                    End If
                    If dblCost <> 0 Then
                        dblRebate = dblRebate + dblCost * (1 + dblRate) * dblCurTime
                        dbl数量 = dbl数量 - dblCurTime
                    End If
                    rsTmp.MoveNext
                Next
                If dbl数量 <> 0 Then GoTo EndCalc '数量未分解完毕,库存不足,不打折
                dblCurMoney = dblRebate * (1 + dbl加班加价率)
                If dblCurMoney < dblMinMoney Then strMin费别 = rsBase!费别: dblMinMoney = dblCurMoney
            End If
        End If
    End If
    
EndCalc:
    If dblMinMoney <> 9999999999999# Then
        str费别 = strMin费别
        ActualMoney = Format(dblMinMoney, gstrDec)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
        IIF(lng病人科室ID <> 0, " And A.ID<>[2]", "") & _
        IIF(blnBed, " And Exists(Select 科室ID From 床位状况记录 Where 科室ID=A.ID)", "") & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", "," & strTmp & ",", lng病人科室ID)
    
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

Public Function GetAuditName(ByVal strName As String) As String
'功能：从"实习医生/审核医生"中取审核医生名
    GetAuditName = Mid(strName, InStr(strName, "/") + 1)
End Function

Public Function HaveAuditPriv(Optional ByVal str姓名 As String) As Boolean
'功能：判断当前人员是否具有"执业医师"的资格
'参数：str姓名=判断指定人员，不传入时判断当前人员
    Dim rsTmp As New ADODB.Recordset
    Dim strHave As String, strNone As String
    Dim strSQL As String
    
    If str姓名 = "" Then str姓名 = UserInfo.姓名
    
    If InStr(strHave & "|", "|" & str姓名 & "|") > 0 Then
        HaveAuditPriv = True: Exit Function
    ElseIf InStr(strNone & "|", "|" & str姓名 & "|") > 0 Then
        HaveAuditPriv = False: Exit Function
    End If
    
    On Error GoTo errH
    strSQL = "Select B.分类 From 人员表 A,执业类别 B Where A.姓名=[1] And A.执业类别=B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "HaveAuditPriv", str姓名)
    If Not rsTmp.EOF Then
        If rsTmp!分类 = "执业医师" Or rsTmp!分类 = "执业助理医师" Then HaveAuditPriv = True
    End If
    If HaveAuditPriv Then
        strHave = strHave & "|" & str姓名
    Else
        strNone = strNone & "|" & str姓名
    End If
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str性质)
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
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID)
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
        " Select M.病人ID,M.主页ID,M.病人科室ID,M.执行科室ID,C.ID,C.类别,C.是否变价,D.跟踪在用,B.收入项目ID,A.数量," & _
        " Nvl(A.从项,0) as 从项,C.加班加价,B.加班加价率,B.附术收费率,Decode(Nvl(C.是否变价,0),1,A.单价,B.现价) as 单价" & _
        " From 病人医嘱记录 M,病人医嘱计价 A,收费价目 B,收费项目目录 C,材料特性 D" & _
        " Where A.收费细目ID=B.收费细目ID And B.收费细目ID=C.ID" & _
        " And C.ID=D.材料ID(+) And M.ID=A.医嘱ID And M.ID=[1]" & _
        " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
        " Order by 从项"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID)
    If gbln从项汇总折扣 And Not rsTmp.EOF And str费别 <> "" Then
        rsTmp.Filter = "从项=1"
        If Not rsTmp.EOF Then blnHaveSub = True
        rsTmp.Filter = 0
    End If
    For i = 1 To rsTmp.RecordCount
        If InStr(",5,6,7,", rsTmp!类别) > 0 And Nvl(rsTmp!是否变价, 0) = 1 Then
            '设定的计价中时价药品单价计算
            lng执行科室ID = Get收费执行科室ID(rsTmp!病人ID, Nvl(rsTmp!主页ID, 0), rsTmp!类别, rsTmp!ID, 4, Nvl(rsTmp!病人科室ID, 0), 0, IIF(Not IsNull(rsTmp!主页ID), 2, 1))
            dbl单价 = Format(CalcDrugPrice(rsTmp!ID, lng执行科室ID, dbl数量 * Nvl(rsTmp!数量, 0), , True), "0.00000")
        ElseIf rsTmp!类别 = "4" And Nvl(rsTmp!是否变价, 0) = 1 And Nvl(rsTmp!跟踪在用, 0) = 1 Then
            '设定的计价中时价卫材单价计算
            lng执行科室ID = Get收费执行科室ID(rsTmp!病人ID, Nvl(rsTmp!主页ID, 0), rsTmp!类别, rsTmp!ID, 4, Nvl(rsTmp!病人科室ID, 0), 0, IIF(Not IsNull(rsTmp!主页ID), 2, 1), Nvl(rsTmp!执行科室ID, 0))
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
    
    strTmp = IIF(int范围 = 1, "门诊", "住院")
    
    strSQL = _
        " Select 药品ID,Sum(Nvl(可用数量,0)) as 库存 From 药品库存" & _
        " Where (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate))" & _
        " And 性质 = 1 And 库房ID=[1]" & IIF(lng药品ID <> 0, " And 药品ID=[2]", "") & _
        " Group by 药品ID Having Sum(Nvl(可用数量,0))<>0"
    strSQL = "Select A.药品ID,A.剂量系数,A." & strTmp & "包装,A." & strTmp & "单位,A.可否分零," & _
        " A.药房分批,B.是否变价,C.库存/A." & strTmp & "包装 as 库存,B.编码,Nvl(D.名称,B.名称) as 名称,B.规格,B.产地,B.撤档时间,B.服务对象" & _
        " From 药品规格 A,收费项目目录 B,(" & strSQL & ") C,收费项目别名 D" & _
        " Where A.药品ID=B.ID And A.药品ID=C.药品ID(+)" & _
        " And B.ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=[5]" & _
        IIF(bln停用, " And B.服务对象 IN([3],3) And (B.撤档时间 is NULL Or B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))", "") & _
        " And A.药名ID=[4]" & IIF(lng药品ID <> 0, " And A.药品ID=[2]", "") & _
        " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng药房ID, lng药品ID, int范围, lng药名ID, IIF(gbln商品名, 3, 1))
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
    On Error GoTo ErrHand

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
            vntNo = IIF(IsNull(!最大号码), 0, !最大号码)
            
            strSQL = "Select Nvl(Max(病人ID),0)+1 as 病人ID From 病人信息 Where 病人ID>=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
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
            .Update "最大号码", IIF(vntNo - 10 > 0, vntNo - 10, 1)
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
                blnByDate = (IIF(IsNull(!参数值), 1, !参数值) = 2)
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
            vntNo = IIF(IsNull(!最大号码), 0, !最大号码)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(住院号),0)+1 as 住院号 From 病人信息 Where 住院号>=[1]"
            Else
                strSQL = "Select Nvl(Max(住院号),To_Number(To_Char(Sysdate,'YYMM')||'0000'))+1 as 住院号" & _
                    " From 病人信息 Where 住院号 Like To_Number(To_Char(Sysdate,'YYMM'))||'%' And 住院号>=[1]"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
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
                .Update "最大号码", IIF(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "最大号码", IIF(vntNo - 10 > Val(Format(curDate, "YYMM0000")), vntNo - 10, Val(Format(curDate, "YYMM0001")))
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
                blnByDate = (IIF(IsNull(!参数值), 1, !参数值) = 2)
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
            vntNo = IIF(IsNull(!最大号码), 0, !最大号码)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(门诊号),0)+1 as 门诊号 From 病人信息 Where 门诊号>=[1]"
            Else
                strSQL = "Select Nvl(Max(门诊号),To_Number(To_Char(Sysdate,'YYMMDD')||'0000'))+1 as 门诊号" & _
                    " From 病人信息 Where 门诊号 Like To_Number(To_Char(Sysdate,'YYMMDD'))||'%' And 门诊号>=[1]"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", Val(vntNo))
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
                .Update "最大号码", IIF(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "最大号码", IIF(vntNo - 10 > Val(Format(curDate, "YYMMDD0000")), vntNo - 10, Val(Format(curDate, "YYMMDD0001")))
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
            
            vntNo = Val(IIF(IsNull(!最大号码), 0, !最大号码)) + 1
            
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
            strYear = IIF(intYear < 10, CStr(intYear), Chr(55 + intYear))
            vntNo = IIF(IsNull(!最大号码), "", !最大号码)
            
            If IIF(IsNull(!编号规则), 0, !编号规则) = 1 Then
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
ErrHand:
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!类别ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDayMoney(lng病人ID As Long) As Currency
'功能：获取指定病人当天发生的费用总额
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng病人ID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID, int场合)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng科室ID)
    DeptIsWoman = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitExe(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer) As String
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
        " Where A.病人ID=[1] And Nvl(A.主页ID,0)=[2] And Nvl(A.婴儿,0)=[3]" & _
        " And B.医嘱ID=A.ID And B.执行部门ID+0 IN(" & strSQL & ")" & _
        " And B.执行状态 IN(0,3) And A.诊疗项目ID=C.ID And B.执行部门ID=D.ID" & _
        " And Not (A.诊疗类别 IN('F','G','D') And A.相关ID is Not NULL)" & _
        " And Not (A.诊疗类别='Z' And Nvl(C.操作类型,'0')<>'0')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng病人ID, lng主页ID, int婴儿)
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

Public Function ExistWaitDrug(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer) As String
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
        " And A.病人ID=[1] And Nvl(A.主页ID,0)=[2] And Nvl(A.婴儿费,0)=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng病人ID, lng主页ID, int婴儿)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID)
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

Public Function ItemIsVarPrice(ByVal lng项目ID As Long) As Boolean
'功能：判断指定项目是否变价(非药品和跟踪在用的卫材)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select A.类别,A.是否变价,B.跟踪在用 From 收费项目目录 A,材料特性 B Where A.ID=B.材料ID(+) And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str编码, lng项目ID)
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
    On Error GoTo ErrHand
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        zlGetSymbol = IIF(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

ErrHand:
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, lng发送号)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO, int记录性质)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO)

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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng号别ID)
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng号别ID, str号别)
        If Not rsTmp.EOF Then GetRegistRoom = rsTmp!门诊诊室
    ElseIf int分诊 = 3 Then
        '平均分诊：当前分配=1表示下次应取的当前诊室
        strSQL = "Select * From 挂号安排诊室 Where 号表ID=" & lng号别ID
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlPublic", adOpenStatic, adLockOptimistic) '可写记录集
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngID)
        
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
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DeptExist(ByVal str工作性质 As String, ByVal int服务对象 As Integer) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 部门ID From 部门性质说明 Where 工作性质=[1] And 服务对象 IN([2],3) And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str工作性质, int服务对象)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
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
    
    strSQL = "Select NO From H" & strTable & " Where NO=[1] And Rownum<2" & IIF(strWhere <> "", " And " & strWhere, "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO)
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
        IIF(lng发送号 <> 0, " And A.发送号+0=[2]", "") & _
        " And A.医嘱ID IN(" & strSQL & ")" & _
        " Group by B.NO"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, lng发送号)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str名称)
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", str医生)
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
        strMsg = """" & str医嘱 & """要求的处方职务不满足：" & vbCrLf & vbCrLf & IIF(bln医保, "对医保或公费病人,", "") & _
            "该药品要求职务至少为""" & Split(STR_职务, ",")(int职务B - 1) & """才能下达,而医生""" & str医生 & """未设置职务。"
    ElseIf int职务B < int职务A Then
        '数值越小职务越高
        strMsg = """" & str医嘱 & """要求的处方职务不满足：" & vbCrLf & vbCrLf & IIF(bln医保, "对医保或公费病人,", "") & _
            "该药品要求职务至少为""" & Split(STR_职务, ",")(int职务B - 1) & """才能下达,而医生""" & str医生 & """的职务为""" & Split(STR_职务, ",")(int职务A - 1) & """。"
    End If
    CheckOneDuty = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiHaveOut(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
'功能：判断病人是否已出院(包括预出院),用于并发操作判断
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 病人ID From 病案主页 Where (出院日期 is Not Null Or Nvl(状态,0)=3) And 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then PatiHaveOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadAdviceSignSource(ByVal int操作类型 As Integer, _
    ByVal lng病人ID As Long, ByVal varTime As Variant, strIDs As String, _
    ByVal lng签名ID As Long, ByVal blnMoved As Boolean, strSource As String, _
    Optional ByVal lng前提ID As Long, Optional ByVal colStopTime As Collection) As Integer
'功能：获取病人用于电子签名/验证的医嘱源文内容
'参数：
'  int操作类型=要签名/验证签名的医嘱状态
'  签名时传入：
'    lng病人ID
'    varTime=病人挂号单号或主页ID
'    strIDs=指定要签名的医嘱ID序列(组ID)
'    lng前提ID=新开医嘱要签名的医嘱来源(是否医技)
'    colStopTime=停止医嘱签名时，传入包含医嘱执行终止时间的数据
'  验证签名时：
'    lng签名ID=签名记录的ID
'    blnMoved=是否医嘱数据已转出
'返回：签名/验证签名的源文生成规则
'      strIDs=签名/验证签名的医嘱ID序列(每个明细ID)
'      strSource=签名/验证签名的医嘱源文
    Dim rsTmp As New ADODB.Recordset
    Dim str组IDs As String, strSQL As String, i As Long
    Dim arrField As Variant, strField As String
    Dim strLine As String, intRule As Integer
    
    On Error GoTo errH
    
    str组IDs = strIDs
    strSource = "": strIDs = ""
    intRule = 1 '这是最新的医嘱签名源文生成规则编号
    
    If lng签名ID = 0 Then
        '签名时
        If int操作类型 = 1 Then
            '对新开的医嘱进行签名：本次就诊/住院当前医生新下达的未签名医嘱
            strSQL = _
                " Select A.* From 病人医嘱记录 A,病人医嘱状态 B" & _
                " Where A.ID=B.医嘱ID And B.签名ID is Null And B.操作类型=1" & _
                " And A.医嘱状态=1 And A.开嘱医生=[3] And Nvl(A.前提ID,0)=[5] And A.病人ID=[1]" & _
                IIF(TypeName(varTime) = "String", " And A.挂号单=[2]", " And A.主页ID=[2]") & _
                IIF(str组IDs <> "", " And Instr([4],','||Nvl(A.相关ID,A.ID)||',')>0", "") & _
                " Order by A.婴儿,Nvl(A.相关ID,A.ID),A.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng病人ID, varTime, UserInfo.姓名, "," & str组IDs & ",", lng前提ID)
        Else
            '对要作废或停止的医嘱进行签名：新开时签了名的指定医嘱，不一定是当前医生下达
            strSQL = _
                " Select A.* From 病人医嘱记录 A,病人医嘱状态 B" & _
                " Where A.ID=B.医嘱ID And B.签名ID is Not Null And B.操作类型=1 And A.病人ID=[1]" & _
                IIF(TypeName(varTime) = "String", " And A.挂号单=[2]", " And A.主页ID=[2]") & _
                IIF(str组IDs <> "", " And Instr([3],','||Nvl(A.相关ID,A.ID)||',')>0", "") & _
                " Order by A.婴儿,Nvl(A.相关ID,A.ID),A.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng病人ID, varTime, "," & str组IDs & ",")
        End If
    Else
        '验证签名时:先读取签名时的源文生成规则
        strSQL = "Select 签名规则 From 医嘱签名记录 Where ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "医嘱签名记录", "H医嘱签名记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng签名ID)
        If Not rsTmp.EOF Then intRule = Nvl(rsTmp!签名规则, 1)
        '--
        strSQL = _
            " Select A.* From 病人医嘱记录 A,病人医嘱状态 B" & _
                " Where A.ID=B.医嘱ID And B.签名ID=[1] Order by A.婴儿,Nvl(A.相关ID,A.ID),A.序号"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng签名ID)
    End If
    
    '医嘱源文的不同生成规则
    If intRule = 1 Then
        If int操作类型 = 8 Then
            strField = "ID,相关ID,姓名,性别,年龄,婴儿,医嘱期效,开始执行时间,医嘱内容,标本部位,单次用量,总给予量," & _
                "医生嘱托,执行频次,频率次数,频率间隔,间隔单位,执行时间方案,执行终止时间,执行性质,紧急标志,开嘱医生,开嘱时间"
        Else
            strField = "ID,相关ID,姓名,性别,年龄,婴儿,医嘱期效,开始执行时间,医嘱内容,标本部位,单次用量,总给予量," & _
                "医生嘱托,执行频次,频率次数,频率间隔,间隔单位,执行时间方案,执行性质,紧急标志,开嘱医生,开嘱时间"
        End If
    End If
    arrField = Split(strField, ",")
        
    '生成医嘱签名源文
    Do While Not rsTmp.EOF
        strLine = ""
        For i = 0 To UBound(arrField)
            If lng签名ID = 0 And int操作类型 = 8 And arrField(i) = "执行终止时间" Then
                '停止医嘱签名时,对终止时间特殊处理：由于是在执行过程之前取签名源文,这时还未写入数据库
                strLine = strLine & vbTab & colStopTime("_" & Nvl(rsTmp!相关ID, rsTmp!ID))
            Else
                If IsNull(rsTmp.Fields(arrField(i)).Value) Then
                    strLine = strLine & vbTab & ""
                Else
                    If rsTmp.Fields(arrField(i)).Type = adDBTimeStamp Then
                        strLine = strLine & vbTab & Format(rsTmp.Fields(arrField(i)).Value, "yyyy-MM-dd HH:mm:ss")
                    Else
                        strLine = strLine & vbTab & rsTmp.Fields(arrField(i)).Value
                    End If
                End If
            End If
        Next
        strSource = strSource & vbCrLf & Mid(strLine, 2)
        strIDs = strIDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    strSource = Mid(strSource, 3)
    strIDs = Mid(strIDs, 2)
    
    ReadAdviceSignSource = intRule
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistsDiagNoses(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int类型 As Integer) As Boolean
'功能：检查病人指定的诊断是否存在
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 记录来源,疾病ID,诊断ID,诊断描述,是否疑诊 From 病人诊断记录" & _
        " Where 病人ID=[1] And Nvl(主页ID,0)=[2] And 诊断类型=[3] And 取消时间 Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng病人ID, lng主页ID, int类型)
    ExistsDiagNoses = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistsSpecAdvice(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：判断病人本次住院是否已经下达了确认的特殊医嘱(出院，转院，死亡)
'返回：如果存在，返回医嘱提示信息。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    strSQL = "Select A.姓名,A.婴儿,A.医嘱内容 From 病人医嘱记录 A,诊疗项目目录 B" & _
        " Where A.诊疗项目ID=B.ID And A.诊疗类别='Z' And B.操作类型 IN('5','6','11')" & _
        " And A.医嘱状态 Not IN(1,2,4) And A.病人ID=[1] And A.主页ID=[2]" & _
        " Order by A.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng病人ID, lng主页ID)
    Do While Not rsTmp.EOF
        strMsg = strMsg & vbCrLf & "　　●" & rsTmp!医嘱内容 & IIF(Nvl(rsTmp!婴儿, 0) <> 0, "(婴儿" & Nvl(rsTmp!婴儿, 0) & ")", "")
        rsTmp.MoveNext
    Loop
    If strMsg <> "" Then
        rsTmp.MoveFirst
        strMsg = "提醒您，病人""" & rsTmp!姓名 & """已经确认下达以下特殊医嘱：" & vbCrLf & strMsg & vbCrLf & ""
    End If
    ExistsSpecAdvice = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check适用用法(ByVal lng用法ID As Long, ByVal lng项目ID As Long, ByVal int来源 As Integer) As Boolean
'功能：检查指定的用法是否适用于指定的项目
'参数：int来源=1-门诊,2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Count(A.用法ID) as 总数,Max(Decode(A.用法ID,[2],1,0)) as 指定" & _
        " From 诊疗用法用量 A,诊疗项目目录 B" & _
        " Where A.用法ID=B.ID And B.服务对象 IN([3],3) And A.项目ID=[1] And A.性质>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID, lng用法ID, int来源)
    If Nvl(rsTmp!总数, 0) <= 1 Then
        Check适用用法 = True
    ElseIf Nvl(rsTmp!指定, 0) = 1 Then
        Check适用用法 = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceTime(ByVal lng医嘱ID As Long, ByVal int类型 As Integer) As Date
'功能：读取医嘱指定操作的时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    strSQL = "Select Max(操作时间) as 时间 From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, int类型)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!时间) Then
            GetAdviceTime = rsTmp!时间
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetAdviceSign(ByVal lng医嘱ID As Long, ByVal int类型 As Integer, ByVal str人员 As String, ByVal dat时间 As Date) As Long
'功能：获取指定医嘱操作的签名ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 签名ID From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=[2] And 操作人员=[3] And 操作时间=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, int类型, str人员, dat时间)
    If Not rsTmp.EOF Then
        GetAdviceSign = Nvl(rsTmp!签名ID, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
