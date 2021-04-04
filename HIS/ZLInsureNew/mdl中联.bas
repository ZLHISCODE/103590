Attribute VB_Name = "mdl中联"
Option Explicit

Public Function 医保初始化_中联(ByVal int险类 As Integer) As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false

    
    '为了避免授权难度增加，此处不再进行对各个医保表数据的检查
    医保初始化_中联 = True
End Function

Public Function 身份标识_中联(Optional bytType As Byte, Optional lng病人ID As Long, Optional ByVal int险类 As Integer) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    
    Dim strTmpIden As String
    
    strTmpIden = frmIdentify中联.ShowCard(bytType, lng病人ID, int险类)
    身份标识_中联 = strTmpIden
End Function

Public Function 个人余额_中联(ByVal lng病人ID As Long, ByVal intinsure As Integer) As Currency
'功能: 提取参保病人个人帐户余额
'参数: bytYear-余额类型,0-所有余额,1-本年余额,2-往年余额
'返回: 返回个人帐户余额的金额
    Dim rsTemp As New ADODB.Recordset

    
    gstrSQL = "select A.帐户余额 from 保险帐户 A where A.病人ID=[1] and A.险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", lng病人ID, intinsure)
    
    If rsTemp.EOF Then
        个人余额_中联 = 0
    Else
        个人余额_中联 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
    End If

End Function

Public Function 门诊结算_中联(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, _
            ByVal cur全自费 As Currency, ByVal cur首先自付 As Currency, ByVal intinsure As Integer, Optional ByRef strAdvance As String = "") As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    Dim cur票据总金额 As Currency
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date
On Error GoTo ErrH
'    '此时所有收费细目必然有对应的医保编码
'    gstrSQL = "Select 病人ID,结帐金额  From 病人费用记录 Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9"
'    Call OpenRecordset(rsTemp, "模拟医保")
'
'    Do Until rsTemp.EOF
'        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
'
'        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
'        rsTemp.MoveNext
'    Loop
    
    '---------------------------------------------------------------------------------------------
    '填写结算表
    curDate = zlDatabase.Currentdate
    
    '帐户年度信息
    Call Get帐户信息(intinsure, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    '参数:
    '   性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN
    '   累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,发生费用金额_IN,全自付金额_IN,首先自付金额_IN
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,支付顺序号_IN,主页ID_IN,
    '   中途结帐_IN,备注_IN
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & g结算数据.发生费用金额 & "," & g结算数据.全自费金额 & "," & g结算数据.首先自付金额 & "," & _
        g结算数据.进入统筹金额 & "," & g结算数据.统筹报销金额 & ",0,0," & cur个人帐户 & "," & g结算数据.优惠金额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    门诊结算_中联 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function


Public Function 门诊结算冲销_中联(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, ByVal intinsure As Integer, Optional ByRef strAdvance As String = "") As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset
    Dim rs退费 As New ADODB.Recordset
    Dim lng冲销ID As Long
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur票据总金额 As Currency, cur全自费 As Currency, cur首先自付 As Currency, cur进入统筹 As Currency
    Dim curDate As Date
On Error GoTo ErrH
        
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,发生费用金额,全自付金额,首先自付金额,进入统筹金额  From 保险结算记录 Where 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", lng结帐ID)
        
    lng病人ID = rsTemp("病人ID")
        
    cur票据总金额 = IIf(IsNull(rsTemp("发生费用金额")), 0, rsTemp("发生费用金额")) * -1
    cur全自费 = IIf(IsNull(rsTemp("全自付金额")), 0, rsTemp("全自付金额")) * -1
    cur首先自付 = IIf(IsNull(rsTemp("首先自付金额")), 0, rsTemp("首先自付金额")) * -1
    cur进入统筹 = IIf(IsNull(rsTemp("进入统筹金额")), 0, rsTemp("进入统筹金额")) * -1
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rs退费 = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", lng结帐ID)
    
    lng冲销ID = rs退费("结帐ID")
    
    '帐户年度信息
    Call Get帐户信息(intinsure, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & intinsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 & "," & cur全自费 & "," & cur首先自付 & "," & _
        cur进入统筹 & ",0,0,0," & cur个人帐户 * -1 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")

    门诊结算冲销_中联 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 个人帐户转预交_中联(lng预交ID As Long, cur个人帐户 As Currency, lng病人ID As Long, ByVal intinsure As Integer) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date
    
    '---------------------------------------------------------------------------------------------
    '填写结算表
    curDate = zlDatabase.Currentdate
    
    '帐户年度信息
    Call Get帐户信息(intinsure, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(3," & lng预交ID & "," & intinsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur个人帐户 & ",0,0,0,0,0,0," & _
        cur个人帐户 & ",0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    个人帐户转预交_中联 = True
End Function


Public Function 个人帐户转预交冲销_中联(lng预交ID As Long, cur个人帐户 As Currency, lng病人ID As Long, ByVal intinsure As Integer) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    Dim rs退费 As New ADODB.Recordset
    Dim lng冲销ID As Long
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur票据总金额 As Currency
    Dim curDate As Date
        
        
    curDate = zlDatabase.Currentdate
    
    '退费
    gstrSQL = "select distinct A.ID from 病人预交记录 A,病人预交记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", lng预交ID)
    
    lng冲销ID = rsTemp("ID")
    
    '帐户年度信息
    Call Get帐户信息(intinsure, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & intinsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur个人帐户 * -1 & ",0,0,0,0,0,0," & _
        cur个人帐户 * -1 & ",0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")

    个人帐户转预交冲销_中联 = True
    
End Function

Public Function 入院登记_中联(lng病人ID As Long, ByVal intinsure As Integer) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false

    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    入院登记_中联 = True
End Function

Public Function 出院登记_中联(lng病人ID As Long, lng主页ID As Long, ByVal intinsure As Integer) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    出院登记_中联 = True
End Function

Public Function 出院登记撤消_中联(lng病人ID As Long, lng主页ID As Long, ByVal intinsure As Integer) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
            '取入院登记验证所返回的顺序号
    On Error GoTo errHandle
    
        
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    出院登记撤消_中联 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_中联(rs费用明细 As Recordset, ByVal intinsure As Integer) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '结算要求：NO、序号、病人ID、医保项目编码、收费类别、收费名称、开单部门、规格、产地、数量、价格、金额、医生,登记时间(发生时间),婴儿费,保险大类ID
    Dim rs大类汇总 As Recordset     '按医保支付大类汇总得到
    Dim rs算法 As New ADODB.Recordset          '保存
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng中心 As Long
    Dim lng在职 As Long, lng年龄段 As Long, lng年龄 As Long
    Dim dblTemp As Double, lng档次 As Long
    
    Dim dbl最大金额  As Double ''对一个按住院日计算的项目，最多能得到的金额
    Dim dbl已报销金额 As Double, dbl累计进入 As Double
    Dim dbl下限 As Double, dbl上限 As Double, dbl分段进入 As Double, dbl分段报销 As Double
    
    Dim cls医保 As New clsInsure
    Dim bln个人帐户支付全自费 As Boolean, bln个人帐户支付首先自付 As Boolean, bln个人帐户支付超限 As Boolean
    Dim cur全自费 As Currency, cur首先自付 As Currency
    Dim bln全额统筹 As Boolean, bln无起付线 As Boolean, bln无封顶线 As Boolean
    
    Dim bln跨年结算 As Boolean   '对于自贡医保，如果是跨年结算，即使该病人是第二次结帐。各分段计算也是从头开始
    Dim dbl多次起付线和 As Double, dbl多次进入统筹和 As Double   '多次是指该病人以前结帐的累计
    Dim dbl计算起付线 As Double, dbl本次起付线 As Double
    Dim dblTemp1 As Double
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '1、初始化一些变量
    Set gcol结算计算 = New Collection
    With g结算数据
        .病人ID = rs费用明细("病人ID")
        
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", CLng(rs费用明细("病人ID")))
        If IsNull(rsTemp("主页ID")) = True Then
            MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        .主页ID = rsTemp("主页ID")
        .年度 = Int(Format(zlDatabase.Currentdate, "yyyy"))
    End With
    
    bln个人帐户支付全自费 = cls医保.GetCapability(support结算帐户全自费, 0, intinsure)
    bln个人帐户支付首先自付 = cls医保.GetCapability(support结算帐户首先自付, 0, intinsure)
    bln个人帐户支付超限 = cls医保.GetCapability(support结算帐户超限, 0, intinsure)

        
    '1.2 读出病人的入院时间
    gstrSQL = "select 入院日期,nvl(出院日期,to_date('3000-01-01','yyyy-MM-dd')) as 出院日期 " & _
              "from 病案主页 where 病人ID=[1] and 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", g结算数据.病人ID, g结算数据.主页ID)
    If rsTemp("出院日期") = CDate("3000-01-01") Then
        g结算数据.中途结帐 = 1
    Else
        '表示该病人已经出院
        g结算数据.中途结帐 = 0
    End If

    '1.3 读出本次住院期间累计结帐情况
    gstrSQL = "select nvl(sum(A.起付线),0) as 起付线,nvl(sum(A.进入统筹金额),0) as 进入统筹金额 " & _
              "  from 保险结算记录 A,病人结帐记录 B " & _
              "  Where A.病人ID = [1] And A.主页ID = [2]" & _
              " And A.险类 = [3] And A.记录ID = B.ID "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", g结算数据.病人ID, g结算数据.主页ID, intinsure)
    dbl多次起付线和 = rsTemp("起付线")
    dbl多次进入统筹和 = rsTemp("进入统筹金额")
    
    With g结算数据
        gstrSQL = "select A.中心,A.人员身份,A.在职,A.年龄段," & _
                  "      B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计" & _
                  " from 保险帐户 A,帐户年度信息 B" & _
                  " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) " & _
                  "     and B.年度(+)=[1] and A.病人ID=[2] and A.险类=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", .年度, .病人ID, intinsure)
        
        lng中心 = IIf(IsNull(rsTemp("中心")), 0, rsTemp("中心"))
        lng在职 = IIf(IsNull(rsTemp("在职")), 1, rsTemp("在职"))
        lng年龄 = IIf(IsNull(rsTemp("年龄段")), 0, rsTemp("年龄段"))
        .住院次数 = IIf(IsNull(rsTemp("住院次数累计")), 0, rsTemp("住院次数累计"))
        .帐户累计增加 = IIf(IsNull(rsTemp("帐户增加累计")), 0, rsTemp("帐户增加累计"))
        .帐户累计支出 = IIf(IsNull(rsTemp("帐户支出累计")), 0, rsTemp("帐户支出累计"))
        .累计进入统筹 = IIf(IsNull(rsTemp("进入统筹累计")), 0, rsTemp("进入统筹累计"))
        .累计统筹报销 = IIf(IsNull(rsTemp("统筹报销累计")), 0, rsTemp("统筹报销累计"))
    
        
        gstrSQL = "select 年龄段,nvl(全额统筹,0) as 全额统筹 ,nvl(无起付线,0) as 无起付线 ,nvl(无封顶线,0) as 无封顶线 " & _
                " from 保险年龄段" & _
                " where 险类=[1] and nvl(中心,0)=[2]" & _
                "       and 在职=[3] and 下限<=[4] and ([4]<=上限 or 上限=0)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", intinsure, lng中心, lng在职, lng年龄)
        If rsTemp.RecordCount = 0 Then
            MsgBox "请在“保险类别管理”中设置年龄段与费用档。", vbInformation, gstrSysName
            Exit Function
        End If
        lng年龄段 = rsTemp("年龄段")
        bln全额统筹 = (rsTemp("全额统筹") = 1)
        bln无起付线 = (rsTemp("无起付线") = 1)
        bln无封顶线 = (rsTemp("无封顶线") = 1)
    End With
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '2、按统筹支付项目合计发生金额和数量
    '2.1、初始化记录集
    Set rs大类汇总 = New ADODB.Recordset
    With rs大类汇总
        If .State = adStateOpen Then .Close
        .Fields.Append "保险大类ID", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adDouble, 8, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .Fields.Append "统筹金额", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With

    Do Until rs费用明细.EOF
    '装数据写入记录集，供其它窗体使用
        If rs费用明细("保险项目否") = 1 Then
            If rs大类汇总.RecordCount = 0 Then
                rs大类汇总.AddNew
                rs大类汇总("保险大类ID") = rs费用明细("保险大类ID")
                rs大类汇总("数量") = rs费用明细("数量")
                rs大类汇总("金额") = rs费用明细("金额")
            Else
                rs大类汇总.MoveFirst
                rs大类汇总.Find "保险大类ID=" & rs费用明细("保险大类ID")
                If rs大类汇总.EOF Then
                    rs大类汇总.AddNew
                    rs大类汇总("保险大类ID") = rs费用明细("保险大类ID")
                    rs大类汇总("数量") = rs费用明细("数量")
                    rs大类汇总("金额") = rs费用明细("金额")
                Else
                    rs大类汇总("数量") = rs大类汇总("数量") + rs费用明细("数量")
                    rs大类汇总("金额") = rs大类汇总("金额") + rs费用明细("金额")
                End If
            End If
            rs大类汇总.Update
        Else
            cur全自费 = cur全自费 + rs费用明细("金额")
        End If
            
        dblTemp = dblTemp + rs费用明细("金额")
        rs费用明细.MoveNext
    Loop
    g结算数据.发生费用金额 = dblTemp
    
    '2.2、计算进入统筹金额
    gstrSQL = "select ID,算法,统筹比额,特准定额,特准天数,是否医保 FROM 保险支付大类  where 险类=[1]"
    Set rs算法 = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", intinsure)
    
    dblTemp = 0
    g结算数据.优惠金额 = 0
    If rs大类汇总.RecordCount > 0 Then rs大类汇总.MoveFirst
    Do Until rs大类汇总.EOF
        
        rs算法.Filter = "ID=" & rs大类汇总("保险大类ID")
        If rs算法.RecordCount > 0 Then
            If rs算法("是否医保") = 1 Then
            
                '算法:1-总额计算项目；2-住院日核定项目
                Select Case Nvl(rs算法!算法, 2)
                Case 1      '1-总额计算项目
                        If rs算法("统筹比额") = 0 Then
                            cur全自费 = cur全自费 + rs大类汇总("金额")
                        Else
                            dblTemp = dblTemp + rs大类汇总("金额") * rs算法("统筹比额") / 100
                        End If
                Case 2      '2-住院日核定项目
                        If Val(rs大类汇总("数量")) > Val(rs算法("特准天数")) Then
                            '如果住院日超过特准天数，那么最大金额就是 特准天数*特准定额 +  (数量-特准天数)*统筹比额
                            '当特准定额或特准天数任一个为0时，就相当于不要特准天数
                            dbl最大金额 = rs算法("特准定额") * rs算法("特准天数") + _
                                (rs大类汇总("数量") - IIf(rs算法("特准定额") = 0 Or rs算法("特准天数") = 0, 0, rs算法("特准天数"))) * rs算法("统筹比额")
                        Else
                            '如果住院日低于特准天数，那么最大金额就是 数量*特准定额 或者 数量*统筹比额
                            '当特准定额或特准天数任一个为0时，就相当于不要特准定额
                            If rs算法("特准定额") = 0 Or rs算法("特准天数") = 0 Then
                                dbl最大金额 = rs大类汇总("数量") * rs算法("统筹比额")
                            Else
                                dbl最大金额 = rs大类汇总("数量") * rs算法("特准定额")
                            End If
                        End If
                        
                        '总金额比最大金额小，就取全部金额；否则只最大金额
                        dblTemp = dblTemp + IIf(rs大类汇总("金额") < dbl最大金额, rs大类汇总("金额"), dbl最大金额)
                        
                        If rs大类汇总("金额") > dbl最大金额 Then
                            '全部算作全自费
                            cur全自费 = cur全自费 + rs大类汇总("金额") - dbl最大金额
                        End If
                Case Else   '3-费用档次计算法
                    If Nvl(rs大类汇总!金额, 0) = 0 Then
                    Else
                        dblTemp1 = 获取费用档次额_中联(Nvl(rs大类汇总!保险大类id, 0), Nvl(rs大类汇总!金额, 0))
                        dblTemp = dblTemp + dblTemp1
                        g结算数据.优惠金额 = g结算数据.优惠金额 + (Nvl(rs大类汇总!金额, 0) - dblTemp1)
                    End If
                End Select
            Else
                cur全自费 = cur全自费 + rs大类汇总("金额")
            End If
        Else
            cur全自费 = cur全自费 + rs大类汇总("金额")
        End If
        rs大类汇总.MoveNext
    Loop
    g结算数据.进入统筹金额 = dblTemp
    g结算数据.全自费金额 = cur全自费
    g结算数据.首先自付金额 = g结算数据.发生费用金额 - cur全自费 - dblTemp - g结算数据.优惠金额
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '3、获得起付线、封顶线、支付比例等数据
    '3.1、获得起付线、封顶线
    With g结算数据
        
        gstrSQL = "select max(decode(A.性质,'A',A.金额,0)) as 封项线 ,max(decode(A.性质,'1',A.金额,0)) as 起付线 " & _
                  "         ,max(decode(A.性质,'" & (.住院次数 + 1) & "',A.金额,0)) as 实际起付线,min(A.金额) as 最低起付线 " & _
                  "  from 保险支付限额 A " & _
                  "  where A.险类=[1] and A.中心=[2] and A.年度=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", intinsure, lng中心, .年度)
                
        If bln无起付线 Then
            .实际起付线 = 0
            .起付线 = 0
        Else
            .起付线 = IIf(IsNull(rsTemp("实际起付线")), 0, rsTemp("实际起付线"))
            If .起付线 = 0 Then
                '一般都会有，如果实在超过了住院次数，就取最后一次（也就是金额最小的一次）
                .起付线 = IIf(IsNull(rsTemp("最低起付线")), 0, rsTemp("最低起付线"))
            End If
            If .起付线 = 0 Then
                MsgBox "请在“年度结算规则”中设置本年度的起付线。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If bln无封顶线 Then
            .封顶线 = 0
        Else
            .封顶线 = IIf(IsNull(rsTemp("封项线")), 0, rsTemp("封项线"))
            If .封顶线 = 0 Then
                MsgBox "请在“年度结算规则”中设置本年度的封顶线。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
        '3.2、根据以前扣除的起付线金额，得出本次的实际起付线
        If dbl多次起付线和 > 0 Then
            '表明该病人肯定有多次结帐
            
            If dbl多次起付线和 > dbl多次进入统筹和 Then
                '该病人的本次结算还要扣除一部分起付线金额
                dbl计算起付线 = dbl多次起付线和 - dbl多次进入统筹和
            Else
                '起付线已经扣完
                dbl计算起付线 = 0
            End If
            
            If .起付线 > dbl多次起付线和 Then
                '调高了起付线，要补这段差值
                .起付线 = .起付线 - dbl多次起付线和
            Else
                '以前的起付线金额已经全额保存，本次不用再保存了
                .起付线 = 0
            End If
                
            dbl计算起付线 = dbl计算起付线 + .起付线
        Else
            dbl计算起付线 = .起付线
        End If
        dbl本次起付线 = dbl计算起付线
    End With
    
    '3.3、取得费用档次
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "select B.档次,B.下限,B.上限,A.比例 " & _
              "  from 保险支付比例 A,保险费用档 B " & _
              "  Where A.险类 =[1] And A.中心 =[2] And A.年度 =[3] And A.在职 =[4] And A.年龄段 =[5]" & _
              "       and A.险类=B.险类 and A.中心=b.中心 and A.档次=B.档次 " & _
              "  order by B.档次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", intinsure, lng中心, g结算数据.年度, lng在职, lng年龄段)
    If rsTemp.RecordCount = 0 Then
        MsgBox "请在“年度结算规则”中设置本年度的统筹支付比例。。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '4、计算该次结算可报销的金额
    dbl累计进入 = 0   '保存分段累计进入统筹
    dbl已报销金额 = g结算数据.累计统筹报销
    g结算数据.统筹报销金额 = 0
    
    If bln跨年结算 = True Then
        '跨年结算就不用考虑以前的结算金额
        dbl多次进入统筹和 = 0
    End If
    Do Until rsTemp.EOF
        dbl分段进入 = 0
        dbl分段报销 = 0
        
        If dbl已报销金额 < g结算数据.封顶线 Or g结算数据.封顶线 = 0 Then    '未超过封顶线或无封顶线
            '还可以继续报销
            dbl下限 = IIf(IsNull(rsTemp("下限")), 0, rsTemp("下限"))
            dbl上限 = IIf(IsNull(rsTemp("上限")), 0, rsTemp("上限"))
            If dbl下限 = 0 Then
                If g结算数据.起付线 > dbl上限 Then
                    MsgBox "该病人的实际起付线比第一档费用的上限还多，请检查保险费用档。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            If g结算数据.进入统筹金额 + dbl多次进入统筹和 > dbl下限 And (dbl多次进入统筹和 < dbl上限 Or dbl上限 = 0) Then
                '该段以前还未计算完全，求出本段需要另外扣除的金额
                dblTemp = 0
                If dbl多次进入统筹和 > dbl下限 Then
                    '以前已经计算过的
                    dblTemp = dbl多次进入统筹和 - dbl下限
                End If
                
                '由于要扣除一部分起付线和已结金额，所以下限金额会有变化
                If dbl下限 + dblTemp + dbl计算起付线 > dbl上限 And dbl上限 > 0 Then
                    dbl下限 = dbl上限
                    dbl计算起付线 = dbl计算起付线 - (dbl上限 - dbl下限 - dblTemp) '本段已经扣完，留着下段扣
                Else
                    dbl下限 = dbl下限 + dbl计算起付线 + dblTemp
                    dbl计算起付线 = 0
                End If
                
                If g结算数据.进入统筹金额 + dbl多次进入统筹和 <= dbl上限 Or dbl上限 = 0 Then
                    '按实际值进入
                    dbl分段进入 = g结算数据.进入统筹金额 + dbl多次进入统筹和 - dbl下限
                    
                    '如果由于加上起付线、或以前的结帐金额，导致进入统筹的金额还不能达到下限，那只能取0
                    If dbl分段进入 < 0 Then dbl分段进入 = 0
                Else
                    '全额进入
                    dbl分段进入 = dbl上限 - dbl下限
                End If
                '按比例求出该段的报销金额
                dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
                dbl分段报销 = Val(Format(dbl分段进入 * rsTemp("比例") / 100, "0.00"))
                
                If dbl已报销金额 + dbl分段报销 > g结算数据.封顶线 And g结算数据.封顶线 <> 0 Then
                    '报销金额超过了封顶线，并且存在封顶线限制
                    dbl分段报销 = g结算数据.封顶线 - dbl已报销金额
                    
                    '倒推进入统筹金额
                    If rsTemp("比例") <> 0 Then
                        dbl分段进入 = dbl分段报销 * 100 / rsTemp("比例")
                    Else
                        dbl分段进入 = 0
                    End If
                End If
                
                '进行格式化
                dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
                dbl分段报销 = Val(Format(dbl分段报销, "0.00"))
                
                dbl已报销金额 = dbl已报销金额 + dbl分段报销
                g结算数据.统筹报销金额 = g结算数据.统筹报销金额 + dbl分段报销
            End If
        End If
        
        '档次、进入统筹金额、统筹报销金额、比例
        lng档次 = IIf(IsNull(rsTemp("档次")), 0, rsTemp("档次"))
        dblTemp = IIf(IsNull(rsTemp("比例")), 0, rsTemp("比例"))
        dbl累计进入 = dbl分段进入 + dbl累计进入
            
        gcol结算计算.Add Array(lng档次, dbl分段进入, dbl分段报销, dblTemp)
        rsTemp.MoveNext
    Loop
    
    g结算数据.实际起付线 = dbl本次起付线 - dbl计算起付线
    
    With g结算数据
        '计算超限自付部分
        .超限自付金额 = .进入统筹金额 - dbl本次起付线 - dbl累计进入
        If .超限自付金额 < 0 Then .超限自付金额 = 0                   '如果进入统筹金额还不到起付线，为负数
    End With
    
    If bln全额统筹 = True Then
        住院虚拟结算_中联 = "医保基金;" & g结算数据.统筹报销金额 + g结算数据.首先自付金额 & ";0"
    Else
        住院虚拟结算_中联 = "医保基金;" & g结算数据.统筹报销金额 & ";0"
    End If
    
    '还需要考虑个人帐户的支付范围
    With g结算数据
        dblTemp = 0   '暂时保存可使用的个人帐户余额
        
        If bln个人帐户支付全自费 = True Then
            dblTemp = dblTemp + .全自费金额
        End If
        
        If bln个人帐户支付首先自付 = True And bln全额统筹 = False Then
            dblTemp = dblTemp + .首先自付金额
        End If
        
        If bln个人帐户支付超限 = True Then
            '只能支付进入统筹，但未报销的部分
            dblTemp = dblTemp + .进入统筹金额 - .统筹报销金额
        Else
            dblTemp = dblTemp + .进入统筹金额 - .统筹报销金额 - .超限自付金额 - .优惠金额
        End If
        '个人帐户支付额不能超过帐户余额，这导致了部分地区病人的帐户余额为零，结算出来仍然有帐户支付额
        If dblTemp > (g结算数据.帐户累计增加 - g结算数据.帐户累计支出) Then
            dblTemp = (g结算数据.帐户累计增加 - g结算数据.帐户累计支出)
        End If
        If .优惠金额 <> 0 Then
            住院虚拟结算_中联 = 住院虚拟结算_中联 & "|优惠金额;" & .优惠金额 & ";0" & "|个人帐户;" & dblTemp & ";1"
        Else
            住院虚拟结算_中联 = 住院虚拟结算_中联 & "|个人帐户;" & dblTemp & ";1"
        End If
    End With
End Function

Public Function 住院结算_中联(lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    Dim rsTemp As New ADODB.Recordset
    Dim cur个人帐户 As Currency
    Dim var结算计算 As Variant
On Error GoTo ErrH
    '求个人帐户支付金额
    gstrSQL = "Select Nvl(冲预交,0) as 金额 From 病人预交记录 Where 结算方式='个人帐户' And 记录性质=2 And 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", lng结帐ID)
    
    If rsTemp.RecordCount > 0 Then
        cur个人帐户 = rsTemp("金额")
    End If
    
    With g结算数据
        gstrSQL = "zl_帐户年度信息_insert(" & .病人ID & "," & intinsure & "," & .年度 & "," & _
            .帐户累计增加 & "," & .帐户累计支出 + cur个人帐户 & "," & .累计进入统筹 + .进入统筹金额 & "," & _
            .累计统筹报销 + .统筹报销金额 & "," & .住院次数 + 1 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
        
        '参数
        '   性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN
        '   累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,发生费用金额_IN,全自付金额_IN,首先自付金额_IN
        '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,支付顺序号_IN,主页ID_IN,
        '   中途结帐_IN,备注_IN

        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & intinsure & "," & .病人ID & "," & _
            .年度 & "," & .帐户累计增加 & "," & .帐户累计支出 & "," & .累计进入统筹 & "," & _
            .累计统筹报销 & "," & .住院次数 + 1 & "," & .起付线 & "," & .封顶线 & "," & .实际起付线 & "," & _
            .发生费用金额 & "," & .全自费金额 & "," & .首先自付金额 & "," & .进入统筹金额 & "," & .统筹报销金额 & ",0," & _
            .超限自付金额 & "," & cur个人帐户 & "," & g结算数据.优惠金额 & "," & .主页ID & "," & .中途结帐 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
        
        
        For Each var结算计算 In gcol结算计算
            '依次为档次、进入统筹金额、统筹报销金额、比例
            gstrSQL = "zl_保险结算计算_Insert(" & lng结帐ID & "," & _
                var结算计算(0) & "," & var结算计算(1) & "," & var结算计算(2) & "," & var结算计算(3) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
        Next
    End With
    
    住院结算_中联 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_中联(lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rs帐户 As New ADODB.Recordset, rs结算计算 As New ADODB.Recordset
    Dim lng冲销ID As Long
    Dim lng住院次数 As Long, cur帐户增加 As Currency, cur帐户支出 As Currency, cur累计进入统筹 As Currency, cur累计统筹报销 As Currency
    
On Error GoTo ErrH

    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", lng结帐ID)
    
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "Select * " & _
              "  From 保险结算记录 Where 性质=2 and 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", lng结帐ID)
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "该病人的医保结算数据丢失，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    If Can住院结算冲销(rsTemp("病人ID"), rsTemp("主页ID")) = False Then Exit Function
    
    gstrSQL = "select B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计 " & _
              " from 保险帐户 A,帐户年度信息 B " & _
              " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) and B.年度(+)=[1] and A.病人ID=[2] and A.险类=[3]"
    Set rs帐户 = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", Year(zlDatabase.Currentdate), CLng(rsTemp("病人ID")), intinsure)
    
    If rs帐户.EOF = False Then
        lng住院次数 = IIf(IsNull(rs帐户("住院次数累计")), 0, rs帐户("住院次数累计"))
        cur帐户增加 = IIf(IsNull(rs帐户("帐户增加累计")), 0, rs帐户("帐户增加累计"))
        cur帐户支出 = IIf(IsNull(rs帐户("帐户支出累计")), 0, rs帐户("帐户支出累计"))
        cur累计进入统筹 = IIf(IsNull(rs帐户("进入统筹累计")), 0, rs帐户("进入统筹累计"))
        cur累计统筹报销 = IIf(IsNull(rs帐户("统筹报销累计")), 0, rs帐户("统筹报销累计"))
    End If
    
    '将此处的数据保存与主程序的数据保存想成一个事务
    '因此就不需要单独的事务控制
    gstrSQL = "zl_帐户年度信息_insert(" & rsTemp("病人ID") & "," & intinsure & "," & rsTemp("年度") & "," & _
        cur帐户增加 & "," & cur帐户支出 - rsTemp("个人帐户支付") & "," & cur累计进入统筹 - rsTemp("进入统筹金额") & "," & _
        cur累计统筹报销 - rsTemp("统筹报销金额") & "," & lng住院次数 - 1 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    '冲销单据，处理了几个累计
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & intinsure & "," & rsTemp("病人ID") & "," & _
        rsTemp("年度") & "," & cur帐户增加 & "," & cur帐户支出 - rsTemp("个人帐户支付") & "," & cur累计进入统筹 - rsTemp("进入统筹金额") & "," & _
        cur累计统筹报销 - rsTemp("统筹报销金额") & "," & lng住院次数 & "," & rsTemp("起付线") * -1 & "," & rsTemp("封顶线") & "," & rsTemp("实际起付线") * -1 & "," & _
        rsTemp("发生费用金额") * -1 & "," & rsTemp("全自付金额") * -1 & "," & rsTemp("首先自付金额") * -1 & "," & rsTemp("进入统筹金额") * -1 & "," & _
        rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") * -1 & "," & rsTemp("个人帐户支付") * -1 & ",''," & _
        IIf(IsNull(rsTemp("主页ID")), "null", rsTemp("主页ID")) & "," & rsTemp("中途结帐") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    
    gstrSQL = "select 档次,进入统筹金额,统筹报销金额,比例 from 保险结算计算 where 结帐ID=[1]"
    Set rs结算计算 = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", lng结帐ID)
    
    Do Until rs结算计算.EOF
        '依次为档次、进入统筹金额、统筹报销金额、比例
        gstrSQL = "zl_保险结算计算_Insert(" & lng冲销ID & "," & _
            rs结算计算("档次") & "," & rs结算计算("进入统筹金额") * -1 & "," & rs结算计算("统筹报销金额") * -1 & "," & rs结算计算("比例") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
        
        rs结算计算.MoveNext
    Loop
    
    住院结算冲销_中联 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 错误信息_中联(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息

End Function

Public Function 获取费用档次额_中联(ByVal lng大类ID As Long, ByVal dbl金额 As Double) As Double
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:根据大类ID和金额求出进入统筹金额
    '--入参数:
    '--出参数:
    '--返  回:进入统筹金额
    '--编 制:刘兴宏 20040617
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "" & _
        "   Select " & _
        "       nvl(Sum(decode(sign(" & dbl金额 & "-nvl(上限,0)),-1,0, " & _
        "              decode(sign(" & dbl金额 & "-nvl(下限,90009000900099.99)),1, " & _
        "                     decode(nvl(下限,0),0," & dbl金额 & "-nvl(上限,0),下限-nvl(上限,0)),   " & _
        "                      " & dbl金额 & "-nvl(上限,0)))*比例/100)," & dbl金额 & ") as 统筹额  " & _
        "   From 大类档次比例 " & _
        "   Where 大类id=" & lng大类ID

    zlDatabase.OpenRecordset rsTemp, gstrSQL, "计算保险大类的统筹额"
    获取费用档次额_中联 = Nvl(rsTemp!统筹额, dbl金额)
    rsTemp.Close
End Function

Public Function 医保项目_中联(lngItemID As Long, Optional ByVal intinsure As Integer) As String
    Dim rsTemp As New ADODB.Recordset
    

    gstrSQL = "Select nvl(B.名称,'') as 名称 From 保险支付项目 A,保险支付大类 B Where A.险类=" & intinsure & " and A.大类id=B.id and " & "A.收费细目id=" & lngItemID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取模拟医保项目的费用类型"
        
    If rsTemp.RecordCount = 0 Then
        医保项目_中联 = ""
        Exit Function
    End If

    医保项目_中联 = rsTemp!名称
End Function
