Attribute VB_Name = "mdl咸阳"
Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'-------------变量定义
Public gcn咸阳 As New ADODB.Connection        '连接到医保前置服务器

'-------------函数定义

Public Function 医保初始化_咸阳() As Boolean
'功能：建立医保前置机的连接
'返回：初始化成功，返回true；否则，返回false
    
    医保初始化_咸阳 = 检查医保服务器_咸阳
End Function

Public Function 身份标识_咸阳(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：strSelfNO-个人编号，刷卡得到；strSelfPwd-病人密码；bytType-识别类型，0-门诊，1-住院
'返回： 空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim str医保号 As String, rsTemp As New ADODB.Recordset
    Dim STR姓名 As String, str性别 As String, str身份证号码 As String, lng年龄 As Long
    Dim str出生日期 As String, str单位编码 As String
    Dim strIdentify As String, str附加 As String, strComputer As String
    
    Dim cur个人帐户 As Currency
   
    On Error GoTo errHandle
    
    '从前置服务器中读出已经验证身份的病人信息
    If bytType = 0 Then
        '门诊信息
        strComputer = Get机器名
'        strComputer = "work1"
        gstrSQL = "SELECT CardNo AS 医保卡号, UnitNo AS 单位编码, SelfSerial AS 个人序号,PatiName As 姓名 " & _
                   "     ,PatiSex As 性别, 0 AS 年龄, '汉族' As 民族, IdentityCard As 身份证号,balance as 个人帐户余额 " & _
                   " From Outpatients " & _
                   "WHERE Terminal = '" & strComputer & "' AND AcceptTime IS NULL"
    ElseIf bytType = 1 Then
        '住院信息
        gstrSQL = "SELECT CardNo AS 医保卡号, '' AS 单位编码, StaySerial As 个人序号,PatiName AS 姓名" & _
                  "       ,PatiSex AS 性别, PatiYear AS 年龄, PatiFolk As 民族, IdentityCard As 身份证号 " & _
                  "  From Inpatients " & _
                   "WHERE AcceptTime IS NULL"
    Else
        '不支持
        Exit Function
    End If
    
    rsTemp.Open gstrSQL, gcn咸阳, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount = 0 Then
        MsgBox "未发现已验证病人信息。", vbInformation, gstrSysName
        Exit Function
    ElseIf rsTemp.RecordCount > 1 Then
        '多于一个时要进行选择
        If frmListSel.ShowSelect(TYPE_咸阳市, rsTemp, "医保卡号", "医保病人选择", "请选择已验证的病人，然后点击确定。") = False Then
            Exit Function
        End If
    End If
    
    str医保号 = rsTemp("医保卡号")
    
    If bytType = 0 Then
        cur个人帐户 = rsTemp("个人帐户余额")
    Else
        cur个人帐户 = 0
    End If
    
    STR姓名 = rsTemp("姓名")
    str性别 = IIf(IsNull(rsTemp("性别")), "", rsTemp("性别"))
    str身份证号码 = IIf(IsNull(rsTemp("身份证号")), "", rsTemp("身份证号"))
    str出生日期 = Get出生日期(str身份证号码, 0)
    If IsDate(str出生日期) Then
        lng年龄 = DateDiff("yyyy", CDate(str出生日期), zlDatabase.Currentdate)
        str出生日期 = Format(CDate(str出生日期), "yyyy-MM-dd")
    Else
        lng年龄 = IIf(IsNull(rsTemp("年龄")), 0, Val(rsTemp("年龄")))
        str出生日期 = Format(DateAdd("yyyy", -1 * lng年龄, zlDatabase.Currentdate), "yyyy-MM-dd") '从年龄倒算生日
    End If
    
    str单位编码 = IIf(IsNull(rsTemp("单位编码")), "", rsTemp("单位编码"))
    
    
    '卡号;医保号;密码;姓名;性别;出生日期;身份证;工作单位
    strIdentify = str医保号 & ";" & str医保号 & ";;" & STR姓名 & ";" & str性别 & ";" & str出生日期 & ";" & str身份证号码 & ";(" & str单位编码 & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    str附加 = ";"                                       '8.中心代码
    str附加 = str附加 & ";" & IIf(IsNull(rsTemp("个人序号")), "", rsTemp("个人序号"))             '9.顺序号
    str附加 = str附加 & ";"                             '10人员身份
    str附加 = str附加 & ";" & cur个人帐户               '11帐户余额
    str附加 = str附加 & ";0"                            '12当前状态
    str附加 = str附加 & ";"                             '13病种ID
    str附加 = str附加 & ";1"                            '14在职(1,2)
    str附加 = str附加 & ";"                             '15退休证号
    str附加 = str附加 & ";" & lng年龄                   '16年龄段
    str附加 = str附加 & ";"                             '17灰度级
    str附加 = str附加 & ";" & cur个人帐户              '18帐户增加累计
    str附加 = str附加 & ";0"                            '19帐户支出累计
    str附加 = str附加 & ";"                             '20进入统筹累计
    str附加 = str附加 & ";"                             '21统筹报销累计
    str附加 = str附加 & ";"                             '22住院次数累计
    
    lng病人ID = BuildPatiInfo(bytType, strIdentify & str附加, lng病人ID, TYPE_咸阳市)
    
    '返回格式:中间插入病人ID
    If lng病人ID <> 0 Then
        身份标识_咸阳 = strIdentify & ";" & lng病人ID & str附加
        
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_咸阳(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'参数:
'返回: 返回个人帐户余额的金额
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '从数据库中读取（因为刚才才保存了的，应该是准确的）
    gstrSQL = "Select 帐户余额 From 保险帐户 where 险类=[1] and 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "咸阳医保", TYPE_咸阳市, lng病人ID)

    If rsTemp.EOF = False Then
        个人余额_咸阳 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
    Else
        个人余额_咸阳 = 0
    End If
    
    
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_咸阳(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保

    Dim str处方号 As String, rsTemp As New ADODB.Recordset
    Dim str医保号 As String, STR姓名 As String, cur个人帐户 As Currency, dat发生时间 As Date
    
    Dim strmachine As String
    
    On Error GoTo errHandle
    
    strmachine = Get机器名()
    
    
    ''''''''
'    strmachine = "WORK1"
    
    ''''''''''''''''''''''
    
    '获得医保号
    gstrSQL = "select B.医保号,C.姓名 " & _
              "  from 保险帐户 B,病人信息 C " & _
              "  where B.病人id=[1] and B.险类=[2] and B.病人ID=C.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", CLng(rs明细("病人ID")), TYPE_咸阳市)
    str医保号 = rsTemp("医保号")
    STR姓名 = rsTemp("姓名")
   dat发生时间 = zlDatabase.Currentdate
    
    '得到该病人的医保处方号
    str处方号 = InputBox("请输入医保病人专用处方号：", "门诊预算")
    If str处方号 = "" Then
        Exit Function
    End If
            
    If zlCommFun.StrIsValid(str处方号, 7) = False Then
        MsgBox "处方号中含有非法字符。", vbInformation, gstrSysName
        Exit Function
    End If
    If Len(str处方号) < 7 Then
        MsgBox "处方号长度不足7位。", vbInformation, gstrSysName
        Exit Function
    End If
    
    
  
    
    '首先完成费用明细的传输
    
    On Error Resume Next
    gcn咸阳.BeginTrans
    
    '如果是门诊，将接收时间设置，表明该病人得到处理（门诊病人的主键是：医保卡号+接收时间）
    gstrSQL = "UPDATE Outpatients Set AcceptTime = GETDATE() WHERE Terminal='" & strmachine & _
                "' and cardno = '" & str医保号 & "' and accepttime is null "
    gcn咸阳.Execute gstrSQL
    
    gcn咸阳.Execute "Delete from ClinicExses where CardNo='" & str医保号 & "' and AcceptTime is null"
    gcn咸阳.Execute "Delete ClinicSettles Where CardNo='" & str医保号 & "' and AcceptTime is null"
    Do Until rs明细.EOF
        gstrSQL = "select A.编码,A.名称,A.计算单位,A.类别,B.名称 as 大类,B.统筹比额  " & _
                   " from 收费细目 A , " & _
                   " (select B. 收费细目ID,C.名称,C.统筹比额 from 保险支付大类 C,保险支付项目 B where B.险类=" & TYPE_咸阳市 & " and B.大类ID=C.ID) B " & _
                   " where A.ID=B.收费细目ID(+) and A.ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", CLng(rs明细("收费细目ID")))
        If IsNull(rsTemp("大类")) = True Then
            MsgBox "收费项目“" & rsTemp("名称") & "”没有设置对应的保险大类。", vbInformation, gstrSysName
            gcn咸阳.RollbackTrans
            Exit Function
        End If
        
        gstrSQL = "INSERT INTO ClinicExses(ID, RecipeNo, CardNo, PatiName, RecordTime, ItemCode, ItemName, ItemUnit, Price, Amount, Money, ExseKind, InsureKind, PayTax) " & _
                  "VALUES('" & str处方号 & "_" & rs明细.AbsolutePosition & "','" & str处方号 & "','" & str医保号 & "','" & STR姓名 & "','" & _
                  Format(dat发生时间, "yyyy-MM-dd HH:mm:ss") & "','" & ToVarchar(rsTemp("编码"), 10) & "','" & ToVarchar(rsTemp("名称"), 90) & "','" & ToVarchar(rsTemp("计算单位"), 6) & "'," & _
                  Format(rs明细("单价"), "0.0000") & "," & Format(rs明细("数量"), "0.000") & "," & Format(rs明细("实收金额"), "0.000") & ",'" & _
                  Switch(rsTemp("类别") = "7", "0", rsTemp("类别") = "5", "1", rsTemp("类别") = "6", "2", rsTemp("类别") = "J", "4", True, 3) & "','" & _
                  IIf(rsTemp("统筹比额") = 100, "0", IIf(rsTemp("统筹比额") = 0, "2", "1")) & "'," & Format(100 - rsTemp("统筹比额"), "0;;\0") & ")"
        gcn咸阳.Execute gstrSQL
        If Err <> 0 Then
            MsgBox "医保费用明细产生失败。", vbInformation, gstrSysName
            gcn咸阳.RollbackTrans
            Exit Function
        End If
        
        rs明细.MoveNext
    Loop
    gcn咸阳.CommitTrans
    
    On Error GoTo errHandle
    '读取前置服务器是否完成结算
    gstrSQL = "SELECT CardPay AS 卡上支付, CashPay AS 现金支付, TotalExse AS 费用合计 " & _
              " FROM ClinicSettles WHERE CardNo = '" & str医保号 & "' AND AcceptTime IS NULL"
    If frm等待返回.WaitForYB(rsTemp, gstrSQL) = False Then Exit Function
    If rsTemp.EOF = True Then
        cur个人帐户 = 0
        Exit Function
    Else
        cur个人帐户 = IIf(IsNull(rsTemp("卡上支付")), 0, rsTemp("卡上支付"))
    End If
    
    '改变前置服务器的门诊结算记录结算时间（门诊结算结果的主键是：医保卡号+接收时间）
    gstrSQL = "UPDATE ClinicSettles Set AcceptTime = GETDATE() WHERE cardno = '" & str医保号 & "' and accepttime is null"
    gcn咸阳.Execute gstrSQL
    
    str结算方式 = "个人帐户;" & cur个人帐户 & ";0"  '允许修改个人帐户
    门诊虚拟结算_咸阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_咸阳(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim str医保号 As String, StrInput As String, arrOutput  As Variant
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset
    Dim str操作员 As String, cur发生费用, datCurr As Date
    
    Dim rs帐户余额 As New Recordset
    Dim cur帐户余额 As Currency
    
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From 门诊费用记录 Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "咸阳医保", lng结帐ID)
    
    If rs明细.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "没有填写收费记录"
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str操作员 = ToVarchar(IIf(IsNull(rs明细("操作员姓名")), UserInfo.姓名, rs明细("操作员姓名")), 20)
    
    Do Until rs明细.EOF
        cur发生费用 = cur发生费用 + rs明细("结帐金额")
        rs明细.MoveNext
    Loop
    
    '调用结算
    
    '保存结算记录
    '---------------------------------------------------------------------------------------------
    datCurr = zlDatabase.Currentdate
    
    
    gstrSQL = "select 帐户余额 from 保险帐户 where 险类=[1] and 病人id=[2]"
    Set rs帐户余额 = zlDatabase.OpenSQLRecord(gstrSQL, "咸阳医保", TYPE_咸阳市, lng病人ID)
    
    If rs帐户余额.EOF Then
        cur帐户余额 = 0
    Else
        cur帐户余额 = rs帐户余额.Fields(0)
    End If
    
    
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_咸阳市 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户余额 & ",0,0,0,0,0,0,0," & cur发生费用 & ",0,0," & _
        "0,0,0,0," & cur个人帐户 & ",'')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "咸阳医保")
    '---------------------------------------------------------------------------------------------

    门诊结算_咸阳 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function


Public Function 门诊结算冲销_咸阳(lng结帐ID As Long, cur个人帐户 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    
    门诊结算冲销_咸阳 = False
End Function

Public Function 个人帐户转预交_咸阳(lng预交ID As Long, curMoney As Currency) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
           
    '由于咸阳医保不支持该业务，所以强行返回失败
    
    个人帐户转预交_咸阳 = False
End Function

Public Function 入院登记_咸阳(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    '获得病人医保号
    gstrSQL = "select 医保号 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "入院登记", lng病人ID, TYPE_咸阳市)
    str医保号 = rsTemp("医保号")
    
    '确认已经入院登记成功
    gstrSQL = "UPDATE Inpatients Set AcceptTime = GETDATE() WHERE cardno = '" & str医保号 & "' and accepttime is null "
    gcn咸阳.Execute gstrSQL
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_咸阳市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "咸阳医保")
    
    入院登记_咸阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_咸阳(ByVal lng病人ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'返回：交易成功返回true；否则，返回false
    
    On Error GoTo errHandle
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_咸阳市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "咸阳医保")
    
    出院登记_咸阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_咸阳(rsExse As Recordset, ByVal lng病人ID As Long, ByVal str医保号 As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合
    '记录性质,记录状态,NO、序号、收费类别、收费名称、开单部门、规格、产地、数量、价格、金额、医生,
    '登记时间,婴儿费,医保项目编码、保险大类ID、保险项目否、是否上传
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    Dim rsTemp As New ADODB.Recordset, rs大类 As New ADODB.Recordset
    Dim STR姓名 As String, dat发生时间 As Date
    Dim cur统筹支付 As Currency, cur个人帐户 As Currency, cur发生费用 As Currency
    Dim blnReturn As Boolean        '是否成功获取医保结算数据
    Dim str项目编码 As String
    Dim str费用类型 As String
    
    On Error GoTo errHandle
    
    '获得医保号
    gstrSQL = "select B.医保号,C.姓名 " & _
              "  from 保险帐户 B,病人信息 C " & _
              "  where B.病人id=[1] and B.险类=[2] and B.病人ID=C.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", lng病人ID, TYPE_咸阳市)
    STR姓名 = rsTemp("姓名")
    dat发生时间 = zlDatabase.Currentdate
    
    '首先完成费用明细的传输
    gstrSQL = "select ID,编码,名称,统筹比额,特准定额 FROM 保险支付大类  where 险类=[1]"
    Set rs大类 = zlDatabase.OpenSQLRecord(gstrSQL, "咸阳医保", TYPE_咸阳市)
    
    On Error Resume Next
    gcn咸阳.BeginTrans
    gcn咸阳.Execute "Delete from InpatiExses where CardNo='" & str医保号 & "' and AcceptTime is null"
    Do Until rsExse.EOF
        cur发生费用 = cur发生费用 + rsExse("金额")
        rs大类.Filter = "ID=" & rsExse("保险大类ID")
        
        If rs大类.EOF = True Then
            MsgBox "收费项目“" & rsExse("收费名称") & "”没有设置对应的保险大类。", vbInformation, gstrSysName
            gcn咸阳.RollbackTrans
            Exit Function
        End If
        
        '判断病人费用记录中是否上传标志是否已为0，如果为0，表示没有上传过，否则，表示已上传过，不再上传
        If rsExse("是否上传") = 0 Then
            '由于对方接收的项目编码是我们系统中收费细目对应的编码，而不是对方医保的编码，这一点是咸阳医保与其他医保的区别之一，所以，这里去取我们系统中收费细目的编码
            gstrSQL = "select distinct 编码 from 收费细目 where 名称=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "咸阳医保", CStr(rsExse("收费名称")))
            
            If rsTemp.EOF Then
                str项目编码 = "001"
            Else
                str项目编码 = ToVarchar(rsTemp.Fields("编码"), 10)
            End If
            
            gstrSQL = "select b.编码 " _
                    & " From " _
                    & " (select 收入项目id from 住院费用记录 " _
                    & "where no =[1] and 记录性质=[2]" & _
                     " and 序号=[3]" & _
                     " and 记录状态=[4]" & _
                     " and 病人id=[5])  a,收入项目 b " & _
                     " Where a.收入项目id = b.ID "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "咸阳医保", CStr(rsExse("NO")), CLng(rsExse("记录性质")), CLng(rsExse("序号")), CLng(rsExse("记录状态")), CLng(rsExse("病人id")))
            
            If rsTemp.EOF Then
                str费用类型 = "001"
            Else
                str费用类型 = ToVarchar(rsTemp.Fields("编码"), 10)
            End If
            
            gstrSQL = rsExse("NO") & "_" & rsExse("序号") & "_" & rsExse("记录性质") & "_" & rsExse("记录状态")
            gstrSQL = "INSERT INTO InpatiExses(ID, CardNo, PatiName, RecordTime, ItemCode, ItemName, ItemUnit, Price, Amount, Money, ExseKind, InsureKind, PayTax) " & _
                      "VALUES('" & gstrSQL & "','" & str医保号 & "','" & STR姓名 & "','" & _
                      Format(rsExse("发生时间"), "yyyy-MM-dd HH:mm:ss") & "','" & str项目编码 & "','" & ToVarchar(rsExse("收费名称"), 90) & "','个'," & _
                      Format(rsExse("价格"), "0.0000") & "," & Format(rsExse("数量"), "0.000") & "," & Format(rsExse("金额"), "0.000") & ",'" & _
                      str费用类型 & "','" & _
                      IIf(rs大类("统筹比额") = 100, "0", IIf(rs大类("统筹比额") = 0, "2", "1")) & "'," & Format(100 - rs大类("统筹比额"), "0;;\0") & ")"
            gcn咸阳.Execute gstrSQL
        End If
        rsExse.MoveNext
        If Err <> 0 Then
            MsgBox "医保费用明细产生失败。", vbInformation, gstrSysName
            gcn咸阳.RollbackTrans
            Exit Function
        End If
    Loop
    gcn咸阳.CommitTrans
    
    On Error GoTo errHandle
    '读取前置服务器是否完成结算
    gstrSQL = "SELECT AgentPay as 统筹支付, CardPay AS 卡上支付, CashPay AS 现金支付, TotalExse AS 费用合计,FlowExse as 超限金额 " & _
              " FROM InpatiSettles WHERE CardNo = '" & str医保号 & "' AND AcceptTime IS NULL"
    blnReturn = frm等待返回.WaitForYB(rsTemp, gstrSQL)
    If blnReturn Then blnReturn = (rsTemp.RecordCount <> 0)
    
    If blnReturn = False Then
        cur个人帐户 = 0
        cur统筹支付 = 0
        With g结算数据
            .病人ID = lng病人ID
            .统筹报销金额 = cur统筹支付
            .个人帐户支付 = cur个人帐户
            .发生费用金额 = cur发生费用
            
            .超限自付金额 = 0
        End With
        Exit Function
    Else
        cur个人帐户 = IIf(IsNull(rsTemp("卡上支付")), 0, rsTemp("卡上支付"))
        cur统筹支付 = IIf(IsNull(rsTemp("统筹支付")), 0, rsTemp("统筹支付"))
        With g结算数据
            .病人ID = lng病人ID
            .统筹报销金额 = cur统筹支付
            .个人帐户支付 = cur个人帐户
            .发生费用金额 = cur发生费用
            
            .超限自付金额 = IIf(IsNull(rsTemp("超限金额")), 0, rsTemp("超限金额"))
        End With
    End If
    
    住院虚拟结算_咸阳 = "医保基金;" & cur统筹支付 & ";0"
    If cur个人帐户 > 0 Then
        住院虚拟结算_咸阳 = 住院虚拟结算_咸阳 & "|个人帐户;" & cur个人帐户 & ";0"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_咸阳(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID     病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset, str医保号 As String
    Dim datCurr As Date
    
    If g结算数据.病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人没有完成医保的预结算操作，不能进行结算。"
        Exit Function
    End If
    
    On Error GoTo errHandle
    '获得医保号
    gstrSQL = "select B.医保号,C.姓名 " & _
              "  from 保险帐户 B,病人信息 C " & _
              "  where B.病人id=[1] and B.险类=[2] and B.病人ID=C.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", lng病人ID, TYPE_咸阳市)
    str医保号 = rsTemp("医保号")
    
    
    '调用结算
    '改变前置服务器的门诊结算记录结算时间
    gstrSQL = "UPDATE InpatiSettles Set AcceptTime = GETDATE() WHERE cardno = '" & str医保号 & "' and accepttime is null"
    gcn咸阳.Execute gstrSQL
    
    
    '更新本地结算的所有费用明细的上传标志
    gstrSQL = "Update 住院费用记录 Set 是否上传=1 Where Nvl(是否上传,0)=0 And 结帐ID=" & lng结帐ID
    gcnOracle.Execute gstrSQL
    
    '填写结算表
    datCurr = zlDatabase.Currentdate
    
    With g结算数据
        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_咸阳市 & "," & lng病人ID & "," & _
            Year(datCurr) & ",0,0,0,0,0,0,NULL,0," & .发生费用金额 & ",0,0," & _
            .统筹报销金额 & "," & .统筹报销金额 & ",0,0," & .个人帐户支付 & ",'',0,0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "咸阳医保")
        
        '保险结算计算
        gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",0," & .统筹报销金额 & "," & .统筹报销金额 & ",NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "咸阳医保")
    End With
        
    住院结算_咸阳 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_咸阳(lng结帐ID As Long) As Boolean
'----------------------------------------------------------------
'功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
'参数：lng结帐ID-需要作废的结帐单ID号；
'返回：交易成功返回true；否则，返回false
'注意：1)主要使用结帐恢复交易和费用删除交易；
'      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
'      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
'----------------------------------------------------------------
    
    住院结算冲销_咸阳 = False
End Function

Public Function 检查医保服务器_咸阳() As Boolean
'功能：测试是否可以连接到前置服务器上
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String, strDatabase As String
    
    '如果连接已经打开，那就不用再测试
    If gcn咸阳.State = adStateOpen Then
        检查医保服务器_咸阳 = True
        Exit Function
    End If
    
    On Error GoTo ErrH
    
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保接口", TYPE_咸阳市)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保服务器"
                strServer = strTemp
            Case "医保数据库"
                strDatabase = strTemp
            Case "医保用户名"
                strUser = strTemp
            Case "医保用户密码"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    gcn咸阳.Open "Provider=SQLOLEDB.1;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & _
                ";Initial Catalog=" & strDatabase & ";Data Source=" & strServer
    If Err <> 0 Then
        Err.Raise 9000, gstrSysName, "医保前置服务器连接失败。"
        Exit Function
    End If
    
    检查医保服务器_咸阳 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Private Function Get机器名() As String
'功能：获得当前的机器名
    Dim STRNAME As String, l As Long
    
    STRNAME = Space(256): l = 256
    
    If GetComputerName(STRNAME, l) <> 0 Then
        Get机器名 = TrimStr(STRNAME)
    End If
End Function


