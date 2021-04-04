Attribute VB_Name = "mdl松藻"
Option Explicit
'===============================================================================
'请参见《松藻医保分析.DOC》了解详细情况，以下为大框架
'退休、退职人员不存在自费段
'离休人员（含建国前老工人）就诊发生的医疗费用全部由统筹医疗基金支付
'
'门诊病人和住院病人的报销规则一致
'   1、首先使用个人帐户支付
'   2、在自费段金额以内的自付（相当于起付线）
'   3、超过自费段的，按比例自付
'       0-5000;         自负10%
'       5000-10000;     自负8%
'       10000-∞;       自负2%
'       以上比例分档累加计算
'
'   优先使用个人帐户支付，按年度累计计算，自费段内自付的金额，
'保存在（保险结算记录.累计进入统筹）中，如果超过自费段，则余下的金额，
'按报销规则的第三点计算
'
'===============================================================================
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _

Private Const str特种病 As String = "'精神病','传染病','职业病','工伤','计划生育','癌症','单位支付','家属病人'"
Private mCur个人自付段 As Currency          '记录本次的个人自付段
Private mCur个人自付段_支付 As Currency     '记录本次实际支付的个人自付段部分金额

Public Function 医保初始化_松藻(ByVal int险类 As Integer) As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false

    
    '为了避免授权难度增加，此处不再进行对各个医保表数据的检查
    医保初始化_松藻 = True
End Function

Public Function 身份标识_松藻(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    
    Dim strTmpIden As String
    
    strTmpIden = frmIdentify松藻.ShowCard(bytType, lng病人ID)
    身份标识_松藻 = strTmpIden
End Function

Public Function 个人余额_松藻(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'参数: bytYear-余额类型,0-所有余额,1-本年余额,2-往年余额
'返回: 返回个人帐户余额的金额
    Dim rsTemp As New ADODB.Recordset

    
    gstrSQL = "select A.帐户余额 from 保险帐户 A where A.病人ID=[1] and A.险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "松藻医保", lng病人ID, TYPE_重庆松藻)
    
    If rsTemp.EOF Then
        个人余额_松藻 = 0
    Else
        个人余额_松藻 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
    End If

End Function

Public Function 入院登记_松藻(lng病人ID As Long) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false

    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆松藻 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
    
    入院登记_松藻 = True
End Function

Public Function 出院登记_松藻(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重庆松藻 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
    
    出院登记_松藻 = True
End Function

Public Function 门诊结算_松藻(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, _
            ByVal cur全自费 As Currency, ByVal cur首先自付 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
    Dim var结算计算 As Variant
On Error GoTo ErrH
    mCur个人自付段 = mCur个人自付段 - mCur个人自付段_支付
    With g结算数据
        cur个人帐户 = .个人帐户支付
        '更新帐户年度信息时，其累计统筹报销始终为零（该字段未用）
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_重庆松藻 & "," & .年度 & "," & _
            .帐户累计增加 & "," & .帐户累计支出 + cur个人帐户 & "," & mCur个人自付段 & "," & _
            .累计统筹报销 + .统筹报销金额 & "," & IIf(.主页ID = 0, "NULL", .主页ID) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
        
        '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
        gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_重庆松藻 & "," & .病人ID & "," & _
            .年度 & "," & .帐户累计增加 & "," & .帐户累计支出 & "," & .累计进入统筹 & "," & _
            .累计统筹报销 & "," & IIf(.主页ID = 0, "NULL", .主页ID) & "," & .起付线 & "," & .封顶线 & "," & .实际起付线 & "," & _
            .发生费用金额 & "," & .全自费金额 & "," & .首先自付金额 & "," & .进入统筹金额 & "," & .统筹报销金额 & ",0," & _
            .超限自付金额 & "," & cur个人帐户 & ",'')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
        
        For Each var结算计算 In gcol结算计算
            '依次为档次、进入统筹金额、统筹报销金额、比例
            gstrSQL = "zl_保险结算计算_Insert(" & lng结帐ID & "," & _
                var结算计算(0) & "," & var结算计算(1) & "," & var结算计算(2) & "," & var结算计算(3) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
        Next
    End With

    门诊结算_松藻 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function


Public Function 门诊结算冲销_松藻(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset
    Dim rs帐户 As New ADODB.Recordset
    Dim rs结算计算 As New ADODB.Recordset
    Dim lng冲销ID As Long
    Dim cur帐户增加 As Currency, cur帐户支出 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim lng住院次数 As Long
    Dim curDate As Date
On Error GoTo ErrH
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "松藻医保")
    
    lng冲销ID = rsTemp("结帐ID")
    
    gstrSQL = "Select * From 保险结算记录 Where 性质=1 And 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "松藻医保")
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "该病人的医保结算数据丢失，不能作废。"
        Exit Function
    End If
    
    '帐户年度信息
    gstrSQL = "select B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计 " & _
              " from 保险帐户 A,帐户年度信息 B " & _
              " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) and B.年度(+)=" & Year(zlDatabase.Currentdate) & " and A.病人ID=" & rsTemp("病人ID") & " and A.险类=" & TYPE_重庆松藻
    Call OpenRecordset(rs帐户, "松藻医保")
    
    If rs帐户.EOF = False Then
        lng住院次数 = IIf(IsNull(rs帐户("住院次数累计")), 0, rs帐户("住院次数累计"))
        cur帐户增加 = IIf(IsNull(rs帐户("帐户增加累计")), 0, rs帐户("帐户增加累计"))
        cur帐户支出 = IIf(IsNull(rs帐户("帐户支出累计")), 0, rs帐户("帐户支出累计"))
        cur进入统筹累计 = IIf(IsNull(rs帐户("进入统筹累计")), 0, rs帐户("进入统筹累计"))
        cur统筹报销累计 = IIf(IsNull(rs帐户("统筹报销累计")), 0, rs帐户("统筹报销累计"))
    End If
            
    gstrSQL = "zl_帐户年度信息_insert(" & rsTemp("病人ID") & "," & TYPE_重庆松藻 & "," & rsTemp("年度") & "," & _
        cur帐户增加 & "," & cur帐户支出 - rsTemp("个人帐户支付") & "," & cur进入统筹累计 & "," & _
        0 & "," & lng住院次数 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_重庆松藻 & "," & rsTemp("病人ID") & "," & _
        rsTemp("年度") & "," & cur帐户增加 & "," & cur帐户支出 - rsTemp("个人帐户支付") & "," & rsTemp("累计进入统筹") * -1 & "," & _
        cur统筹报销累计 & ",NULL," & rsTemp("起付线") * -1 & "," & rsTemp("封顶线") & "," & rsTemp("实际起付线") * -1 & "," & _
        rsTemp("发生费用金额") * -1 & "," & rsTemp("全自付金额") * -1 & "," & rsTemp("首先自付金额") * -1 & "," & rsTemp("进入统筹金额") * -1 & "," & _
        rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") * -1 & "," & rsTemp("个人帐户支付") * -1 & ",NULL)" 'cur统筹报销累计 - rsTemp("统筹报销金额")
    Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
    
    gstrSQL = "select 档次,进入统筹金额,统筹报销金额,比例 from 保险结算计算 where 结帐ID=" & lng结帐ID
    Call OpenRecordset(rs结算计算, "松藻医保")
    
    Do Until rs结算计算.EOF
        '依次为档次、进入统筹金额、统筹报销金额、比例
        gstrSQL = "zl_保险结算计算_Insert(" & lng冲销ID & "," & _
            rs结算计算("档次") & "," & rs结算计算("进入统筹金额") * -1 & "," & rs结算计算("统筹报销金额") * -1 & "," & rs结算计算("比例") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
        
        rs结算计算.MoveNext
    Loop

    门诊结算冲销_松藻 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊虚拟结算_松藻(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."是否允许修改:0-不允许修改;1-允许修改
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '结算要求：NO、序号、病人ID、医保项目编码、收费类别、收费名称、开单部门、规格、产地、数量、价格、金额、医生,登记时间(发生时间),婴儿费,保险大类ID
    Dim rs算法 As New ADODB.Recordset          '保存
    Dim rsTemp As New ADODB.Recordset
    Dim rs大类汇总 As New ADODB.Recordset
    
    Dim lng中心 As Long
    Dim lng在职 As Long, lng年龄段 As Long, lng年龄 As Long
    Dim dblTemp As Double, lng档次 As Long
    
    Dim dbl最大金额  As Double ''对一个按住院日计算的项目，最多能得到的金额
    Dim dbl已报销金额 As Double, dbl累计进入 As Double
    Dim dbl下限 As Double, dbl上限 As Double, dbl分段进入 As Double, dbl分段报销 As Double
    
    Dim cls医保 As New clsInsure
    Dim bln个人帐户支付全自费 As Boolean, bln个人帐户支付首先自付 As Boolean, bln个人帐户支付超限 As Boolean
    Dim cur全自费 As Currency, cur首先自付 As Currency
    Dim bln无起付线 As Boolean, bln无封顶线 As Boolean
    Dim dbl帐户余额
    Dim dbl多次起付线和 As Double   '多次是指该病人以前结帐的累计
    Dim dbl本次起付线 As Double     '本次的起付线
    Dim blnExit As Boolean          '低于个人帐户余额或起付线（自费段），则保存相关记录后退出
    Dim bln离休人员 As Boolean      '离体人员就诊所发生的费用，全由统筹医疗基金支付
    Dim bln特种病患者 As Boolean
    Dim bln精神病传染病患者 As Boolean, bln计划生育 As Boolean, bln家属病人 As Boolean
    Dim str疾病名称 As String, str特准项目 As String, dbl特准项目 As Currency, lng病种ID As Long
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '1、初始化一些变量
    Set gcol结算计算 = New Collection
    With g结算数据
        .病人ID = rs明细("病人ID")
        .主页ID = 0
        .年度 = Int(Format(zlDatabase.Currentdate, "yyyy"))
    End With

    '1.1 读出本次就诊发生的累计结算情况（累计进入统筹大于自费段，则按比例共同负担，但个人帐户仍优先支付）
    '累计进入统筹做为每次支出的自费段金额（门诊取本年度的累计进入统筹金额）
    gstrSQL = "select nvl(sum(A.累计进入统筹),0) as 起付线 " & _
              "  from 保险结算记录 A " & _
              "  Where A.病人ID = " & g结算数据.病人ID & " And A.险类 = " & TYPE_重庆松藻 & " And A.年度=" & g结算数据.年度
    Call OpenRecordset(rsTemp, "虚拟结算")
    dbl多次起付线和 = rsTemp("起付线")
    
    With g结算数据
        g结算数据.统筹报销金额 = 0
        g结算数据.累计进入统筹 = 0
        g结算数据.累计统筹报销 = 0
        g结算数据.全自费金额 = 0
        g结算数据.首先自付金额 = 0
        g结算数据.个人帐户支付 = 0
        
        gstrSQL = "select A.中心,A.人员身份,A.在职,A.年龄段,Nvl(C.类别,0) 病种,Nvl(C.ID,0) 病种ID,C.名称 疾病名称," & _
                  "      B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计" & _
                  " from 保险帐户 A,帐户年度信息 B," & _
                  "         (Select * From 保险病种 Where 类别<>2" & _
                  "          Union " & _
                  "          Select * From 保险病种 Where 类别=2 And 名称 In (" & str特种病 & ")) C" & _
                  " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) And A.险类=C.险类(+) ANd A.病种ID=C.ID(+) " & _
                  "     and B.年度(+)=" & .年度 & " and A.病人ID=" & .病人ID & " and A.险类=" & TYPE_重庆松藻
        Call OpenRecordset(rsTemp, "虚拟结算")
        
        '1-在职;2-退休;3-离休
        '退休及离休人员，不存在起付线（自费段）
        lng中心 = IIf(IsNull(rsTemp("中心")), 0, rsTemp("中心"))
        lng在职 = IIf(IsNull(rsTemp("在职")), 1, rsTemp("在职"))
        lng年龄 = IIf(IsNull(rsTemp("年龄段")), 0, rsTemp("年龄段"))
        bln特种病患者 = (rsTemp!病种 = 2)
        bln离休人员 = (lng在职 = 3)
        lng病种ID = rsTemp!病种ID
        str疾病名称 = IIf(IsNull(rsTemp!疾病名称), "", rsTemp!疾病名称)
        
        .住院次数 = 1   '本医保与住院次数无关
        .帐户累计增加 = IIf(IsNull(rsTemp("帐户增加累计")), 0, rsTemp("帐户增加累计"))
        .帐户累计支出 = IIf(IsNull(rsTemp("帐户支出累计")), 0, rsTemp("帐户支出累计"))
        mCur个人自付段 = IIf(IsNull(rsTemp("进入统筹累计")), 0, rsTemp("进入统筹累计"))
        .累计统筹报销 = IIf(IsNull(rsTemp("统筹报销累计")), 0, rsTemp("统筹报销累计"))
        
        gstrSQL = "select 年龄段,nvl(无起付线,0) as 无起付线,nvl(无封顶线,0) as 无封顶线 " & _
                " from 保险年龄段" & _
                " where 险类=" & TYPE_重庆松藻 & " and nvl(中心,0)=" & lng中心 & _
                "       and 在职=" & lng在职 & " and 下限<=" & lng年龄 & " and (" & lng年龄 & "<=上限 or 上限=0)"
        Call OpenRecordset(rsTemp, "虚拟结算")
        If rsTemp.RecordCount = 0 Then
            MsgBox "请在“保险类别管理”中设置年龄段与费用档。", vbInformation, gstrSysName
            Exit Function
        End If
        lng年龄段 = rsTemp("年龄段")
        bln无起付线 = (rsTemp("无起付线") = 1) Or (lng在职 <> 1)
        bln无封顶线 = (rsTemp("无封顶线") = 1)
    End With
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '2、获取实际发生费用
    If Not cls医保.GetCapability(support允许不设置医保项目, 0, TYPE_重庆松藻) Then
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
    
        Do Until rs明细.EOF
        '装数据写入记录集，供其它窗体使用
            If rs明细("是否医保") = 1 Then
                If rs大类汇总.RecordCount = 0 Then
                    rs大类汇总.AddNew
                    rs大类汇总("保险大类ID") = IIf(IsNull(rs明细("保险支付大类ID")), 0, rs明细("保险支付大类ID"))
                    rs大类汇总("数量") = rs明细("数量")
                    rs大类汇总("金额") = rs明细("实收金额")
                Else
                    rs大类汇总.MoveFirst
                    rs大类汇总.Find "保险大类ID=" & IIf(IsNull(rs明细("保险支付大类ID")), 0, rs明细("保险支付大类ID"))
                    If rs大类汇总.EOF Then
                        rs大类汇总.AddNew
                        rs大类汇总("保险大类ID") = IIf(IsNull(rs明细("保险支付大类ID")), 0, rs明细("保险支付大类ID"))
                        rs大类汇总("数量") = rs明细("数量")
                        rs大类汇总("金额") = rs明细("实收金额")
                    Else
                        rs大类汇总("数量") = rs大类汇总("数量") + rs明细("数量")
                        rs大类汇总("金额") = rs大类汇总("金额") + rs明细("实收金额")
                    End If
                End If
                rs大类汇总.Update
            Else
                cur全自费 = cur全自费 + rs明细("实收金额")
            End If
                
            dblTemp = dblTemp + rs明细("实收金额")
            rs明细.MoveNext
        Loop
        g结算数据.发生费用金额 = dblTemp
        
        '2.2、计算进入统筹金额
        gstrSQL = "select ID,算法,统筹比额,特准定额,特准天数,是否医保 FROM 保险支付大类  where 险类=" & TYPE_重庆松藻
        Call OpenRecordset(rs算法, "松藻医保")
        
        dblTemp = 0
        If rs大类汇总.RecordCount > 0 Then rs大类汇总.MoveFirst
        Do Until rs大类汇总.EOF
            
            rs算法.Filter = "ID=" & rs大类汇总("保险大类ID")
            If rs算法.RecordCount > 0 Then
                If rs算法("是否医保") = 1 Then
                    '算法:1-总额计算项目；2-住院日核定项目
                    If rs算法("算法") = 1 Then
                        If rs算法("统筹比额") = 0 Then
                            cur全自费 = cur全自费 + rs大类汇总("金额")
                        Else
                            dblTemp = dblTemp + rs大类汇总("金额") * rs算法("统筹比额") / 100
                        End If
                    Else
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
                    End If
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
        g结算数据.首先自付金额 = g结算数据.发生费用金额 - cur全自费 - dblTemp
    Else
        '计算费用总额
        If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
        Do Until rs明细.EOF
            dblTemp = dblTemp + rs明细("实收金额")
            rs明细.MoveNext
        Loop
        g结算数据.发生费用金额 = dblTemp
        g结算数据.进入统筹金额 = dblTemp
    End If
    
    '如果是离休人员，就诊发生的所有费用，全部由统筹医疗基金支付，先扣除个人帐户
    If bln离休人员 Then
        dblTemp = 个人余额_松藻(g结算数据.病人ID)
        dblTemp = IIf(dblTemp > 0, dblTemp, 0)
        dblTemp = IIf(g结算数据.进入统筹金额 > dblTemp, dblTemp, g结算数据.进入统筹金额)
        g结算数据.个人帐户支付 = dblTemp
        g结算数据.统筹报销金额 = g结算数据.进入统筹金额 - g结算数据.个人帐户支付
        str结算方式 = "医保基金;" & g结算数据.统筹报销金额 & ";0|个人帐户;" & g结算数据.个人帐户支付 & ";0"
        门诊虚拟结算_松藻 = True
        Exit Function
    End If
    
    '如果是特种病患者
    If bln特种病患者 Then
        
        Dim rs特种病汇总 As New ADODB.Recordset
        str特准项目 = ""
        dbl特准项目 = 0
        bln精神病传染病患者 = (InStr(1, ",精神病,传染病,", "," & str疾病名称 & ",") <> 0)
        bln计划生育 = (InStr(1, ",计划生育,", "," & str疾病名称 & ",") <> 0)
        bln家属病人 = (InStr(1, ",家属病人,", "," & str疾病名称 & ",") <> 0)
        
        If bln家属病人 Then
            '药品费用，由医保基金支付50%
            g结算数据.进入统筹金额 = 0
            With rs明细
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    If InStr(1, ",5,6,7,", "," & !收费类别 & ",") <> 0 And !是否医保 = 1 Then
                        g结算数据.进入统筹金额 = g结算数据.进入统筹金额 + (!实收金额 * 0.5)
                    End If
                    .MoveNext
                Loop
            End With
            g结算数据.统筹报销金额 = g结算数据.进入统筹金额
            str结算方式 = "医保基金;" & g结算数据.统筹报销金额 & ";0|个人帐户;0;0"
            门诊虚拟结算_松藻 = True
            Exit Function
        ElseIf bln计划生育 Then
            str结算方式 = 特种病结算(str疾病名称, lng在职, dbl特准项目, False)
            门诊虚拟结算_松藻 = True
            Exit Function
        Else
            '计算特准项目进入统筹的总额
            With rsTemp
                gstrSQL = "Select 收费细目ID From 保险特准项目 Where 病种ID=" & lng病种ID
                Call OpenRecordset(rsTemp, "虚拟结算")
                
                Do While Not .EOF
                    str特准项目 = str特准项目 & ";" & !收费细目ID
                    .MoveNext
                Loop
                str特准项目 = str特准项目 & ";"
            End With
            
            If Not cls医保.GetCapability(support允许不设置医保项目, 0, TYPE_重庆松藻) Then
                Set rs特种病汇总 = New ADODB.Recordset
                With rs特种病汇总
                    If .State = adStateOpen Then .Close
                    .Fields.Append "保险大类ID", adDouble, 18, adFldIsNullable
                    .Fields.Append "数量", adDouble, 8, adFldIsNullable
                    .Fields.Append "金额", adDouble, 18, adFldIsNullable
                    .Fields.Append "统筹金额", adDouble, 18, adFldIsNullable
                    
                    .CursorLocation = adUseClient
                    .Open , , adOpenStatic, adLockOptimistic
                End With
            
                If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
                Do Until rs明细.EOF
                '装数据写入记录集，供其它窗体使用
                    If rs明细("是否医保") = 1 And InStr(1, str特准项目, ";" & rs明细("收费细目ID") & ";") <> 0 Then
                        If rs特种病汇总.RecordCount = 0 Then
                            rs特种病汇总.AddNew
                            rs特种病汇总("保险大类ID") = IIf(IsNull(rs明细("保险支付大类ID")), 0, rs明细("保险支付大类ID"))
                            rs特种病汇总("数量") = rs明细("数量")
                            rs特种病汇总("金额") = rs明细("实收金额")
                        Else
                            rs特种病汇总.MoveFirst
                            rs特种病汇总.Find "保险大类ID=" & IIf(IsNull(rs明细("保险支付大类ID")), 0, rs明细("保险支付大类ID"))
                            If rs特种病汇总.EOF Then
                                rs特种病汇总.AddNew
                                rs特种病汇总("保险大类ID") = IIf(IsNull(rs明细("保险支付大类ID")), 0, rs明细("保险支付大类ID"))
                                rs特种病汇总("数量") = rs明细("数量")
                                rs特种病汇总("金额") = rs明细("实收金额")
                            Else
                                rs特种病汇总("数量") = rs特种病汇总("数量") + rs明细("数量")
                                rs特种病汇总("金额") = rs特种病汇总("金额") + rs明细("实收金额")
                            End If
                        End If
                        rs特种病汇总.Update
                    End If
                    rs明细.MoveNext
                Loop
                
                '2.2、计算进入统筹金额
                If rs算法.RecordCount <> 0 Then rs算法.MoveFirst
                If rs特种病汇总.RecordCount > 0 Then rs特种病汇总.MoveFirst
                Do Until rs特种病汇总.EOF
                    
                    rs算法.Filter = "ID=" & rs特种病汇总("保险大类ID")
                    If rs算法.RecordCount > 0 Then
                        If rs算法("是否医保") = 1 Then
                            '算法:1-总额计算项目；2-住院日核定项目
                            If rs算法("算法") = 1 Then
                                If rs算法("统筹比额") = 0 Then
                                    cur全自费 = cur全自费 + rs特种病汇总("金额")
                                Else
                                    dbl特准项目 = dbl特准项目 + rs特种病汇总("金额") * rs算法("统筹比额") / 100
                                End If
                            Else
                                If Val(rs特种病汇总("数量")) > Val(rs算法("特准天数")) Then
                                    '如果住院日超过特准天数，那么最大金额就是 特准天数*特准定额 +  (数量-特准天数)*统筹比额
                                    '当特准定额或特准天数任一个为0时，就相当于不要特准天数
                                    dbl最大金额 = rs算法("特准定额") * rs算法("特准天数") + _
                                        (rs特种病汇总("数量") - IIf(rs算法("特准定额") = 0 Or rs算法("特准天数") = 0, 0, rs算法("特准天数"))) * rs算法("统筹比额")
                                Else
                                    '如果住院日低于特准天数，那么最大金额就是 数量*特准定额 或者 数量*统筹比额
                                    '当特准定额或特准天数任一个为0时，就相当于不要特准定额
                                    If rs算法("特准定额") = 0 Or rs算法("特准天数") = 0 Then
                                        dbl最大金额 = rs特种病汇总("数量") * rs算法("统筹比额")
                                    Else
                                        dbl最大金额 = rs特种病汇总("数量") * rs算法("特准定额")
                                    End If
                                End If
                                
                                '总金额比最大金额小，就取全部金额；否则只最大金额
                                dbl特准项目 = dbl特准项目 + IIf(rs特种病汇总("金额") < dbl最大金额, rs特种病汇总("金额"), dbl最大金额)
                            End If
                        End If
                    End If
                    rs特种病汇总.MoveNext
                Loop
            Else
                '计算费用总额
                If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
                Do Until rs明细.EOF
                    If InStr(1, str特准项目, ";" & rs明细("收费细目ID") & ";") <> 0 Then
                        dbl特准项目 = dbl特准项目 + rs明细("实收金额")
                    End If
                    rs明细.MoveNext
                Loop
            End If
            
            '门诊中，精神病、传染病的特准项目由医保基金支付，余下的进入统筹部分，仍按医保规则计算
            If Not bln精神病传染病患者 Then
                str结算方式 = 特种病结算(str疾病名称, lng在职, dbl特准项目, False)
                门诊虚拟结算_松藻 = True
                Exit Function
            Else
                g结算数据.进入统筹金额 = g结算数据.进入统筹金额 - dbl特准项目
                
                '必冲个人帐户
                If bln精神病传染病患者 Then
                    dbl帐户余额 = 个人余额_松藻(g结算数据.病人ID)
                    If dbl帐户余额 > 0 Then
                        If dbl特准项目 <= dbl帐户余额 Then
                            dbl帐户余额 = dbl特准项目
                        End If
                    Else
                        dbl帐户余额 = 0
                    End If
                    g结算数据.个人帐户支付 = dbl帐户余额
                    g结算数据.统筹报销金额 = dbl特准项目 - dbl帐户余额
                Else
                    g结算数据.统筹报销金额 = dbl特准项目
                End If
            End If
        End If
    End If
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '3、减去个人帐户、起付段金额后，剩下的即是进入统筹的金额
    '3.1、获得起付线、封顶线
    With g结算数据
        
        gstrSQL = "select max(decode(A.性质,'A',A.金额,0)) as 封项线 ,max(decode(A.性质,'1',A.金额,0)) as 起付线 " & _
                  "         ,max(decode(A.性质,'" & (.住院次数 + 1) & "',A.金额,0)) as 实际起付线,min(A.金额) as 最低起付线 " & _
                  "  from 保险支付限额 A " & _
                  "  where A.险类=" & TYPE_重庆松藻 & " and A.中心=" & lng中心 & " and A.年度=" & .年度
        Call OpenRecordset(rsTemp, "虚拟结算")
                
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
'            If mCur个人自付段 < .起付线 Then .起付线 = mCur个人自付段
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
    
        '3.2、根据以前扣除的起付线金额（自费段），得出本次的实际起付线
        If dbl多次起付线和 > 0 Then
            '表明该病人肯定有多次结帐
            If .起付线 > dbl多次起付线和 Then
                '调高了起付线，要补这段差值
                .起付线 = .起付线 - dbl多次起付线和
            Else
                '以前的起付线金额已经全额保存，本次不用再保存了
                .起付线 = 0
            End If
                
            dbl本次起付线 = .起付线
        Else
            dbl本次起付线 = .起付线
        End If
    End With
    g结算数据.实际起付线 = dbl本次起付线
    
    '3.3、取得实际进入统筹的金额（先个人帐户支付，再支付本次起付线，余下的进入统筹金额）
    dbl帐户余额 = 个人余额_松藻(g结算数据.病人ID) - g结算数据.个人帐户支付
    '3.3.1.1、使用个人帐户支付，保存支付记录，如果费用低于个人帐户余额则退出
    blnExit = False
    dblTemp = 0
    If dbl帐户余额 >= 0 Then
        If g结算数据.进入统筹金额 <= dbl帐户余额 Then
            dblTemp = g结算数据.进入统筹金额
            blnExit = True
        Else
            dblTemp = dbl帐户余额
        End If
        
        '3.3.1.2、保存个人帐户支付记录
        g结算数据.个人帐户支付 = g结算数据.个人帐户支付 + dblTemp
        str结算方式 = "个人帐户;" & g结算数据.个人帐户支付 & ";0"
        If blnExit Then
            '进入统筹金额保存的是超出自费段部分的金额，因此，本处要清零
            If g结算数据.统筹报销金额 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 = "", "", "|") & "医保基金;" & g结算数据.统筹报销金额 & ";0"
            g结算数据.进入统筹金额 = 0
            门诊虚拟结算_松藻 = True
            Exit Function
        End If
    End If
    g结算数据.进入统筹金额 = g结算数据.进入统筹金额 - dblTemp
    
    '2004-11-25 ZYB
    '减去个人自付段金额，剩下的按比例报销
    blnExit = False
    dblTemp = 0
    If g结算数据.进入统筹金额 >= mCur个人自付段 Then
        g结算数据.进入统筹金额 = g结算数据.进入统筹金额 - mCur个人自付段
        mCur个人自付段_支付 = mCur个人自付段
    Else
        mCur个人自付段_支付 = g结算数据.进入统筹金额
        g结算数据.进入统筹金额 = 0
        blnExit = True
    End If
    If blnExit Then
        '进入统筹金额保存的是超出自费段部分的金额，因此，本处要清零
        If g结算数据.统筹报销金额 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 = "", "", "|") & "医保基金;" & g结算数据.统筹报销金额 & ";0"
        g结算数据.进入统筹金额 = 0
        门诊虚拟结算_松藻 = True
        Exit Function
    End If
    
    '3.3.2.1、在自费段以内的金额，保存累计进入统筹并退出
    blnExit = False
    dblTemp = 0
    If Not bln无起付线 Then
        If dbl本次起付线 > 0 Then
            If g结算数据.进入统筹金额 <= dbl本次起付线 Then
                dblTemp = g结算数据.进入统筹金额
                blnExit = True
            Else
                dblTemp = dbl本次起付线
            End If
            '3.3.2.2、保存累计进入统筹记录
            g结算数据.累计进入统筹 = dblTemp
            If blnExit Then
                '进入统筹金额保存的是超出自费段部分的金额，因此，本处要清零
                If g结算数据.统筹报销金额 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 = "", "", "|") & "医保基金;" & g结算数据.统筹报销金额 & ";0"
                g结算数据.进入统筹金额 = 0
                门诊虚拟结算_松藻 = True
                Exit Function
            End If
        End If
    End If
    
    '----实际进入统筹金额=（实际发生费用金额-个人帐户-自费段剩余金额）
    g结算数据.进入统筹金额 = g结算数据.进入统筹金额 - g结算数据.累计进入统筹
    
    '3.4、取得费用档次
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "select B.档次,B.下限,B.上限,A.比例 " & _
              "  from 保险支付比例 A,保险费用档 B " & _
              "  Where A.险类 =" & TYPE_重庆松藻 & " And A.中心 =" & lng中心 & " And A.年度 =" & g结算数据.年度 & " And A.在职 =" & lng在职 & " And A.年龄段 =" & lng年龄段 & _
              "       and A.险类=B.险类 and A.中心=b.中心 and A.档次=B.档次 " & _
              "  order by B.档次"
    Call OpenRecordset(rsTemp, "虚拟结算")
    If rsTemp.RecordCount = 0 Then
        MsgBox "请在“年度结算规则”中设置本年度的统筹支付比例。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '4、计算该次结算可报销的金额
    dbl累计进入 = 0   '保存分段累计进入统筹
    dbl已报销金额 = g结算数据.累计统筹报销
    
    '待修改 -- dbl多次进入统筹和
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
            
            If g结算数据.进入统筹金额 > dbl下限 Then
                dblTemp = 0
                
                If g结算数据.进入统筹金额 <= dbl上限 Or dbl上限 = 0 Then
                    '按实际值进入
                    dbl分段进入 = g结算数据.进入统筹金额 - dbl下限
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
        
                '档次、进入统筹金额、统筹报销金额、比例
                lng档次 = IIf(IsNull(rsTemp("档次")), 0, rsTemp("档次"))
                dblTemp = IIf(IsNull(rsTemp("比例")), 0, rsTemp("比例"))
                dbl累计进入 = dbl分段进入 + dbl累计进入
                    
                gcol结算计算.Add Array(lng档次, dbl分段进入, dbl分段报销, dblTemp)
            End If
        End If
        rsTemp.MoveNext
    Loop
    str结算方式 = str结算方式 & IIf(str结算方式 = "", "", "|") & "医保基金;" & g结算数据.统筹报销金额 & ";0"
    门诊虚拟结算_松藻 = True
End Function

Public Function 住院虚拟结算_松藻(rs费用明细 As Recordset) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."是否允许修改:0-不允许修改;1-允许修改
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '结算要求：NO、序号、病人ID、医保项目编码、收费类别、收费名称、开单部门、规格、产地、数量、价格、金额、医生,登记时间(发生时间),婴儿费,保险大类ID
    Dim rs算法 As New ADODB.Recordset          '保存
    Dim rsTemp As New ADODB.Recordset
    Dim rs大类汇总 As New ADODB.Recordset
    
    Dim lng中心 As Long
    Dim lng在职 As Long, lng年龄段 As Long, lng年龄 As Long
    Dim dblTemp As Double, lng档次 As Long
    
    Dim dbl最大金额  As Double ''对一个按住院日计算的项目，最多能得到的金额
    Dim dbl已报销金额 As Double, dbl累计进入 As Double
    Dim dbl下限 As Double, dbl上限 As Double, dbl分段进入 As Double, dbl分段报销 As Double
    
    Dim cls医保 As New clsInsure
    Dim bln个人帐户支付全自费 As Boolean, bln个人帐户支付首先自付 As Boolean, bln个人帐户支付超限 As Boolean
    Dim cur全自费 As Currency, cur首先自付 As Currency
    Dim bln无起付线 As Boolean, bln无封顶线 As Boolean
    Dim dbl帐户余额
    Dim dbl多次起付线和 As Double   '多次是指该病人以前结帐的累计
    Dim dbl本次起付线 As Double     '本次的起付线
    Dim blnExit As Boolean          '低于个人帐户余额或起付线（自费段），则保存相关记录后退出
    Dim bln特种病患者 As Boolean    '特种病患者
    Dim bln离休人员 As Boolean, bln家属病人 As Boolean
    Dim str疾病名称 As String, str特准项目 As String, dbl特准项目 As Currency, lng病种ID As Long
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '1、初始化一些变量
    Set gcol结算计算 = New Collection
    With g结算数据
        .病人ID = rs费用明细("病人ID")
        
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=" & rs费用明细("病人ID")
        Call OpenRecordset(rsTemp, "虚拟结算")
        If IsNull(rsTemp("主页ID")) = True Then
            MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        .主页ID = rsTemp("主页ID")
        .年度 = Int(Format(zlDatabase.Currentdate, "yyyy"))
    End With
    
    '1.1 读出病人的入院时间
    gstrSQL = "select 入院日期,nvl(出院日期,to_date('3000-01-01','yyyy-MM-dd')) as 出院日期 " & _
              "from 病案主页 where 病人ID=" & g结算数据.病人ID & " and 主页ID=" & g结算数据.主页ID
    Call OpenRecordset(rsTemp, "虚拟结算")
    If rsTemp("出院日期") = CDate("3000-01-01") Then
        g结算数据.中途结帐 = 1
    Else
        '表示该病人已经出院
        g结算数据.中途结帐 = 0
    End If

    '1.2 读出本次住院期间累计结帐情况（累计进入统筹大于自费段，则按比例共同负担，但个人帐户仍优先支付）
    '累计进入统筹做为每次支出的自费段金额
    gstrSQL = "select nvl(sum(A.累计进入统筹),0) as 起付线 " & _
              "  from 保险结算记录 A " & _
              "  Where A.病人ID = " & g结算数据.病人ID & " And A.险类 = " & TYPE_重庆松藻 & " And A.年度= " & g结算数据.年度
    Call OpenRecordset(rsTemp, "虚拟结算")
    dbl多次起付线和 = rsTemp("起付线")
    
    With g结算数据
        g结算数据.统筹报销金额 = 0
        g结算数据.累计进入统筹 = 0
        g结算数据.累计统筹报销 = 0
        g结算数据.全自费金额 = 0
        g结算数据.首先自付金额 = 0
        g结算数据.个人帐户支付 = 0
        
        gstrSQL = "select A.中心,A.人员身份,A.在职,A.年龄段,Nvl(C.类别,0) 病种,C.名称 疾病名称,Nvl(C.ID,0) 病种ID," & _
                  "      B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计" & _
                  " from 保险帐户 A,帐户年度信息 B," & _
                  "         (Select * From 保险病种 Where 类别<>2" & _
                  "          Union " & _
                  "          Select * From 保险病种 Where 类别=2 And 名称 In (" & str特种病 & ")) C" & _
                  " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) And A.险类=C.险类(+) ANd A.病种ID=C.ID(+) " & _
                  "     and B.年度(+)=" & .年度 & " and A.病人ID=" & .病人ID & " and A.险类=" & TYPE_重庆松藻
        Call OpenRecordset(rsTemp, "虚拟结算")
        
        '1-在职;2-退休;3-离休
        '退休及离休人员，不存在起付线（自费段）
        lng中心 = IIf(IsNull(rsTemp("中心")), 0, rsTemp("中心"))
        lng在职 = IIf(IsNull(rsTemp("在职")), 1, rsTemp("在职"))
        lng年龄 = IIf(IsNull(rsTemp("年龄段")), 0, rsTemp("年龄段"))
        bln特种病患者 = (rsTemp!病种 = 2)
        bln离休人员 = (lng在职 = 3)
        lng病种ID = rsTemp!病种ID
        str疾病名称 = IIf(IsNull(rsTemp!疾病名称), "", rsTemp!疾病名称)
        
        .住院次数 = 1   '本医保与住院次数无关
        .帐户累计增加 = IIf(IsNull(rsTemp("帐户增加累计")), 0, rsTemp("帐户增加累计"))
        .帐户累计支出 = IIf(IsNull(rsTemp("帐户支出累计")), 0, rsTemp("帐户支出累计"))
        '本人自付段，如果小于规定起付线，则起付线取本人自付段
        mCur个人自付段 = IIf(IsNull(rsTemp("进入统筹累计")), 0, rsTemp("进入统筹累计"))
        .累计统筹报销 = IIf(IsNull(rsTemp("统筹报销累计")), 0, rsTemp("统筹报销累计"))
        
        gstrSQL = "select 年龄段,nvl(无起付线,0) as 无起付线,nvl(无封顶线,0) as 无封顶线 " & _
                " from 保险年龄段" & _
                " where 险类=" & TYPE_重庆松藻 & " and nvl(中心,0)=" & lng中心 & _
                "       and 在职=" & lng在职 & " and 下限<=" & lng年龄 & " and (" & lng年龄 & "<=上限 or 上限=0)"
        Call OpenRecordset(rsTemp, "虚拟结算")
        If rsTemp.RecordCount = 0 Then
            MsgBox "请在“保险类别管理”中设置年龄段与费用档。", vbInformation, gstrSysName
            Exit Function
        End If
        lng年龄段 = rsTemp("年龄段")
        bln无起付线 = (rsTemp("无起付线") = 1) Or (lng在职 <> 1)
        bln无封顶线 = (rsTemp("无封顶线") = 1)
    End With
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '2、按统筹支付项目合计发生金额和数量
    '2.1、初始化记录集
    If Not cls医保.GetCapability(support允许不设置医保项目, 0, TYPE_重庆松藻) Then
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
                    rs大类汇总("保险大类ID") = IIf(IsNull(rs费用明细("保险大类ID")), 0, rs费用明细("保险大类ID"))
                    rs大类汇总("数量") = rs费用明细("数量")
                    rs大类汇总("金额") = rs费用明细("金额")
                Else
                    rs大类汇总.MoveFirst
                    rs大类汇总.Find "保险大类ID=" & IIf(IsNull(rs费用明细("保险大类ID")), 0, rs费用明细("保险大类ID"))
                    If rs大类汇总.EOF Then
                        rs大类汇总.AddNew
                        rs大类汇总("保险大类ID") = IIf(IsNull(rs费用明细("保险大类ID")), 0, rs费用明细("保险大类ID"))
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
        gstrSQL = "select ID,算法,统筹比额,特准定额,特准天数,是否医保 FROM 保险支付大类  where 险类=" & TYPE_重庆松藻
        Call OpenRecordset(rs算法, "松藻医保")
        
        dblTemp = 0
        If rs大类汇总.RecordCount > 0 Then rs大类汇总.MoveFirst
        Do Until rs大类汇总.EOF
            
            rs算法.Filter = "ID=" & rs大类汇总("保险大类ID")
            If rs算法.RecordCount > 0 Then
                If rs算法("是否医保") = 1 Then
                    '算法:1-总额计算项目；2-住院日核定项目
                    If rs算法("算法") = 1 Then
                        If rs算法("统筹比额") = 0 Then
                            cur全自费 = cur全自费 + rs大类汇总("金额")
                        Else
                            dblTemp = dblTemp + rs大类汇总("金额") * rs算法("统筹比额") / 100
                        End If
                    Else
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
                    End If
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
        g结算数据.首先自付金额 = g结算数据.发生费用金额 - cur全自费 - dblTemp
    Else
        '计算费用总额
        If rs费用明细.RecordCount <> 0 Then rs费用明细.MoveFirst
        Do Until rs费用明细.EOF
            dblTemp = dblTemp + rs费用明细("金额")
            rs费用明细.MoveNext
        Loop
        g结算数据.发生费用金额 = dblTemp
        g结算数据.进入统筹金额 = dblTemp
    End If
    
    '如果是离休人员，就诊发生的所有费用，扣除个人帐户外，全部由统筹医疗基金支付
    If bln离休人员 Then
        dblTemp = 个人余额_松藻(g结算数据.病人ID)
        dblTemp = IIf(dblTemp > 0, dblTemp, 0)
        dblTemp = IIf(g结算数据.进入统筹金额 > dblTemp, dblTemp, g结算数据.进入统筹金额)
        g结算数据.个人帐户支付 = dblTemp
        g结算数据.统筹报销金额 = g结算数据.进入统筹金额 - dblTemp
        住院虚拟结算_松藻 = "医保基金;" & g结算数据.统筹报销金额 & ";0|个人帐户;" & g结算数据.个人帐户支付 & ";0"
        Exit Function
    End If
    
    '如果是特种病患者
    If bln特种病患者 Then
        
        Dim rs特种病汇总 As New ADODB.Recordset
        str特准项目 = ""
        dbl特准项目 = 0
        bln家属病人 = (InStr(1, ",家属病人,", "," & str疾病名称 & ",") <> 0)
        
        If Not bln家属病人 Then
            住院虚拟结算_松藻 = 特种病结算(str疾病名称, lng在职, 0, True)
            Exit Function
        End If
        
        If bln家属病人 Then
            '药品费、手术费、血费由统筹基金支付50%
            g结算数据.进入统筹金额 = 0
            With rs费用明细
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    If InStr(1, ",5,6,7,F,K", "," & !收费类别 & ",") <> 0 And !保险项目否 = 1 Then
                        g结算数据.进入统筹金额 = g结算数据.进入统筹金额 + (!金额 * 0.5)
                    End If
                    .MoveNext
                Loop
            End With
            g结算数据.统筹报销金额 = g结算数据.进入统筹金额
            住院虚拟结算_松藻 = "医保基金;" & g结算数据.统筹报销金额 & ";0|个人帐户;0;0"
            Exit Function
        Else
            '计算特准项目进入统筹的总额
            With rsTemp
                gstrSQL = "Select 收费细目ID From 保险特准项目 Where 病种ID=" & lng病种ID
                Call OpenRecordset(rsTemp, "虚拟结算")
                
                Do While Not .EOF
                    str特准项目 = str特准项目 & ";" & !收费细目ID
                    .MoveNext
                Loop
                str特准项目 = str特准项目 & ";"
            End With
            
            If Not cls医保.GetCapability(support允许不设置医保项目, 0, TYPE_重庆松藻) Then
                Set rs特种病汇总 = New ADODB.Recordset
                With rs特种病汇总
                    If .State = adStateOpen Then .Close
                    .Fields.Append "保险大类ID", adDouble, 18, adFldIsNullable
                    .Fields.Append "数量", adDouble, 8, adFldIsNullable
                    .Fields.Append "金额", adDouble, 18, adFldIsNullable
                    .Fields.Append "统筹金额", adDouble, 18, adFldIsNullable
                    
                    .CursorLocation = adUseClient
                    .Open , , adOpenStatic, adLockOptimistic
                End With
            
                If rs费用明细.RecordCount <> 0 Then rs费用明细.MoveFirst
                Do Until rs费用明细.EOF
                '装数据写入记录集，供其它窗体使用
                    If rs费用明细("保险项目否") = 1 And InStr(1, str特准项目, ";" & rs费用明细("收费细目ID") & ";") <> 0 Then
                        If rs特种病汇总.RecordCount = 0 Then
                            rs特种病汇总.AddNew
                            rs特种病汇总("保险大类ID") = IIf(IsNull(rs费用明细("保险大类ID")), 0, rs费用明细("保险大类ID"))
                            rs特种病汇总("数量") = rs费用明细("数量")
                            rs特种病汇总("金额") = rs费用明细("金额")
                        Else
                            rs特种病汇总.MoveFirst
                            rs特种病汇总.Find "保险大类ID=" & IIf(IsNull(rs费用明细("保险大类ID")), 0, rs费用明细("保险大类ID"))
                            If rs特种病汇总.EOF Then
                                rs特种病汇总.AddNew
                                rs特种病汇总("保险大类ID") = IIf(IsNull(rs费用明细("保险大类ID")), 0, rs费用明细("保险大类ID"))
                                rs特种病汇总("数量") = rs费用明细("数量")
                                rs特种病汇总("金额") = rs费用明细("金额")
                            Else
                                rs特种病汇总("数量") = rs特种病汇总("数量") + rs费用明细("数量")
                                rs特种病汇总("金额") = rs特种病汇总("金额") + rs费用明细("金额")
                            End If
                        End If
                        rs特种病汇总.Update
                    End If
                    rs费用明细.MoveNext
                Loop
                
                '2.2、计算进入统筹金额
                If rs算法.RecordCount <> 0 Then rs算法.MoveFirst
                If rs特种病汇总.RecordCount > 0 Then rs特种病汇总.MoveFirst
                Do Until rs特种病汇总.EOF
                    
                    rs算法.Filter = "ID=" & rs特种病汇总("保险大类ID")
                    If rs算法.RecordCount > 0 Then
                        If rs算法("是否医保") = 1 Then
                            '算法:1-总额计算项目；2-住院日核定项目
                            If rs算法("算法") = 1 Then
                                If rs算法("统筹比额") = 0 Then
                                    cur全自费 = cur全自费 + rs特种病汇总("金额")
                                Else
                                    dbl特准项目 = dbl特准项目 + rs特种病汇总("金额") * rs算法("统筹比额") / 100
                                End If
                            Else
                                If Val(rs特种病汇总("数量")) > Val(rs算法("特准天数")) Then
                                    '如果住院日超过特准天数，那么最大金额就是 特准天数*特准定额 +  (数量-特准天数)*统筹比额
                                    '当特准定额或特准天数任一个为0时，就相当于不要特准天数
                                    dbl最大金额 = rs算法("特准定额") * rs算法("特准天数") + _
                                        (rs特种病汇总("数量") - IIf(rs算法("特准定额") = 0 Or rs算法("特准天数") = 0, 0, rs算法("特准天数"))) * rs算法("统筹比额")
                                Else
                                    '如果住院日低于特准天数，那么最大金额就是 数量*特准定额 或者 数量*统筹比额
                                    '当特准定额或特准天数任一个为0时，就相当于不要特准定额
                                    If rs算法("特准定额") = 0 Or rs算法("特准天数") = 0 Then
                                        dbl最大金额 = rs特种病汇总("数量") * rs算法("统筹比额")
                                    Else
                                        dbl最大金额 = rs特种病汇总("数量") * rs算法("特准定额")
                                    End If
                                End If
                                
                                '总金额比最大金额小，就取全部金额；否则只最大金额
                                dbl特准项目 = dbl特准项目 + IIf(rs特种病汇总("金额") < dbl最大金额, rs特种病汇总("金额"), dbl最大金额)
                            End If
                        End If
                    End If
                    rs特种病汇总.MoveNext
                Loop
            Else
                '计算费用总额
                If rs费用明细.RecordCount <> 0 Then rs费用明细.MoveFirst
                Do Until rs费用明细.EOF
                    If InStr(1, str特准项目, ";" & rs费用明细("收费细目ID") & ";") <> 0 Then
                        dbl特准项目 = dbl特准项目 + rs费用明细("实收金额")
                    End If
                    rs费用明细.MoveNext
                Loop
            End If
        End If
    End If
    
    '计划生育的特准项目由医保基金支付，余下的进入统筹部分，仍按医保规则计算
    g结算数据.进入统筹金额 = g结算数据.进入统筹金额 - dbl特准项目
    g结算数据.统筹报销金额 = dbl特准项目
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '3、减去个人帐户、起付段金额后，剩下的即是进入统筹的金额
    '3.1、获得起付线、封顶线
    With g结算数据
        
        gstrSQL = "select max(decode(A.性质,'A',A.金额,0)) as 封项线 ,max(decode(A.性质,'1',A.金额,0)) as 起付线 " & _
                  "         ,max(decode(A.性质,'" & (.住院次数 + 1) & "',A.金额,0)) as 实际起付线,min(A.金额) as 最低起付线 " & _
                  "  from 保险支付限额 A " & _
                  "  where A.险类=" & TYPE_重庆松藻 & " and A.中心=" & lng中心 & " and A.年度=" & .年度
        Call OpenRecordset(rsTemp, "虚拟结算")
                
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
'            If mCur个人自付段 < .起付线 Then .起付线 = mCur个人自付段
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
    
        '3.2、根据以前扣除的起付线金额（自费段），得出本次的实际起付线
        If dbl多次起付线和 > 0 Then
            '表明该病人肯定有多次结帐
            If .起付线 > dbl多次起付线和 Then
                '调高了起付线，要补这段差值
                .起付线 = .起付线 - dbl多次起付线和
            Else
                '以前的起付线金额已经全额保存，本次不用再保存了
                .起付线 = 0
            End If
                
            dbl本次起付线 = .起付线
        Else
            dbl本次起付线 = .起付线
        End If
    End With
    g结算数据.实际起付线 = dbl本次起付线
    
    '3.3、取得实际进入统筹的金额（先个人帐户支付，再支付本次起付线，余下的进入统筹金额）
    dbl帐户余额 = 个人余额_松藻(g结算数据.病人ID)
    '3.3.1.1、使用个人帐户支付，保存支付记录，如果费用低于个人帐户余额则退出
    blnExit = False
    dblTemp = 0
    If dbl帐户余额 >= 0 Then
        If g结算数据.进入统筹金额 <= dbl帐户余额 Then
            dblTemp = g结算数据.进入统筹金额
            blnExit = True
        Else
            dblTemp = dbl帐户余额
        End If
        
        '3.3.1.2、保存个人帐户支付记录
        g结算数据.个人帐户支付 = dblTemp
        住院虚拟结算_松藻 = "个人帐户;" & g结算数据.个人帐户支付 & ";0"
        If blnExit Then
            '进入统筹金额保存的是超出自费段部分的金额，因此，本处要清零
            If g结算数据.统筹报销金额 <> 0 Then 住院虚拟结算_松藻 = 住院虚拟结算_松藻 & IIf(住院虚拟结算_松藻 = "", "", "|") & "医保基金;" & g结算数据.统筹报销金额 & ";0"
            g结算数据.进入统筹金额 = 0
            Exit Function
        End If
    End If
    g结算数据.进入统筹金额 = g结算数据.进入统筹金额 - g结算数据.个人帐户支付
    
    '2004-11-25 ZYB
    '减去个人自付段金额，剩下的按比例报销
    blnExit = False
    dblTemp = 0
    If g结算数据.进入统筹金额 >= mCur个人自付段 Then
        g结算数据.进入统筹金额 = g结算数据.进入统筹金额 - mCur个人自付段
        mCur个人自付段_支付 = mCur个人自付段
    Else
        mCur个人自付段_支付 = g结算数据.进入统筹金额
        g结算数据.进入统筹金额 = 0
        blnExit = True
    End If
    If blnExit Then
        '进入统筹金额保存的是超出自费段部分的金额，因此，本处要清零
        If g结算数据.统筹报销金额 <> 0 Then 住院虚拟结算_松藻 = 住院虚拟结算_松藻 & IIf(住院虚拟结算_松藻 = "", "", "|") & "医保基金;" & g结算数据.统筹报销金额 & ";0"
        g结算数据.进入统筹金额 = 0
        Exit Function
    End If
    
    '3.3.2.1、在自费段以内的金额，保存累计进入统筹并退出
    blnExit = False
    dblTemp = 0
    g结算数据.累计进入统筹 = 0
    If Not bln无起付线 Then
        If dbl本次起付线 > 0 Then
            If g结算数据.进入统筹金额 <= dbl本次起付线 Then
                dblTemp = g结算数据.进入统筹金额
                blnExit = True
            Else
                dblTemp = dbl本次起付线
            End If
            '3.3.2.2、保存累计进入统筹记录
            g结算数据.累计进入统筹 = dblTemp
            If blnExit Then
                '进入统筹金额保存的是超出自费段部分的金额，因此，本处要清零
                If g结算数据.统筹报销金额 <> 0 Then 住院虚拟结算_松藻 = 住院虚拟结算_松藻 & IIf(住院虚拟结算_松藻 = "", "", "|") & "医保基金;" & g结算数据.统筹报销金额 & ";0"
                g结算数据.进入统筹金额 = 0
                Exit Function
            End If
        End If
    End If
    
    '----实际进入统筹金额=（实际发生费用金额-个人帐户-自费段剩余金额）
    g结算数据.进入统筹金额 = g结算数据.进入统筹金额 - g结算数据.累计进入统筹
    
    '3.4、取得费用档次
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "select B.档次,B.下限,B.上限,A.比例 " & _
              "  from 保险支付比例 A,保险费用档 B " & _
              "  Where A.险类 =" & TYPE_重庆松藻 & " And A.中心 =" & lng中心 & " And A.年度 =" & g结算数据.年度 & " And A.在职 =" & lng在职 & " And A.年龄段 =" & lng年龄段 & _
              "       and A.险类=B.险类 and A.中心=b.中心 and A.档次=B.档次 " & _
              "  order by B.档次"
    Call OpenRecordset(rsTemp, "虚拟结算")
    If rsTemp.RecordCount = 0 Then
        MsgBox "请在“年度结算规则”中设置本年度的统筹支付比例。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '4、计算该次结算可报销的金额
    dbl累计进入 = 0   '保存分段累计进入统筹
    dbl已报销金额 = g结算数据.累计统筹报销
    
    '待修改 -- dbl多次进入统筹和
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
            
            If g结算数据.进入统筹金额 > dbl下限 Then
                dblTemp = 0
                
                If g结算数据.进入统筹金额 <= dbl上限 Or dbl上限 = 0 Then
                    '按实际值进入
                    dbl分段进入 = g结算数据.进入统筹金额 - dbl下限
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
        
                '档次、进入统筹金额、统筹报销金额、比例
                lng档次 = IIf(IsNull(rsTemp("档次")), 0, rsTemp("档次"))
                dblTemp = IIf(IsNull(rsTemp("比例")), 0, rsTemp("比例"))
                dbl累计进入 = dbl分段进入 + dbl累计进入
                    
                gcol结算计算.Add Array(lng档次, dbl分段进入, dbl分段报销, dblTemp)
            End If
        End If
        rsTemp.MoveNext
    Loop
    住院虚拟结算_松藻 = 住院虚拟结算_松藻 & IIf(住院虚拟结算_松藻 = "", "", "|") & "医保基金;" & g结算数据.统筹报销金额 & ";0"
End Function

Public Function 住院结算_松藻(lng结帐ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    Dim cur个人帐户 As Currency
    Dim var结算计算 As Variant
On Error GoTo ErrH
    With g结算数据
        mCur个人自付段 = mCur个人自付段 - mCur个人自付段_支付
        cur个人帐户 = .个人帐户支付
        gstrSQL = "zl_帐户年度信息_insert(" & .病人ID & "," & TYPE_重庆松藻 & "," & .年度 & "," & _
            .帐户累计增加 & "," & .帐户累计支出 + cur个人帐户 & "," & mCur个人自付段 & "," & _
            .累计统筹报销 + .统筹报销金额 & "," & .主页ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
        
        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_重庆松藻 & "," & .病人ID & "," & _
            .年度 & "," & .帐户累计增加 & "," & .帐户累计支出 & "," & .累计进入统筹 & "," & _
            .累计统筹报销 & "," & .主页ID & "," & .起付线 & "," & .封顶线 & "," & .实际起付线 & "," & _
            .发生费用金额 & "," & .全自费金额 & "," & .首先自付金额 & "," & .进入统筹金额 & "," & .统筹报销金额 & ",0," & _
            .超限自付金额 & "," & cur个人帐户 & ",NULL," & .主页ID & "," & .中途结帐 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
        
        For Each var结算计算 In gcol结算计算
            '依次为档次、进入统筹金额、统筹报销金额、比例
            gstrSQL = "zl_保险结算计算_Insert(" & lng结帐ID & "," & _
                var结算计算(0) & "," & var结算计算(1) & "," & var结算计算(2) & "," & var结算计算(3) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
        Next
    End With
    
    住院结算_松藻 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_松藻(lng结帐ID As Long) As Boolean
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
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "松藻医保")
    
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "Select * " & _
              "  From 保险结算记录 Where 性质=2 and 记录ID='" & lng结帐ID & "'"
    Call OpenRecordset(rsTemp, "松藻医保")
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "该病人的医保结算数据丢失，不能作废。"
        Exit Function
    End If
    
    gstrSQL = "select B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计 " & _
              " from 保险帐户 A,帐户年度信息 B " & _
              " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) and B.年度(+)=" & Year(zlDatabase.Currentdate) & " and A.病人ID=" & rsTemp("病人ID") & " and A.险类=" & TYPE_重庆松藻
    Call OpenRecordset(rs帐户, "松藻医保")
    
    If rs帐户.EOF = False Then
        lng住院次数 = IIf(IsNull(rs帐户("住院次数累计")), 0, rs帐户("住院次数累计"))
        cur帐户增加 = IIf(IsNull(rs帐户("帐户增加累计")), 0, rs帐户("帐户增加累计"))
        cur帐户支出 = IIf(IsNull(rs帐户("帐户支出累计")), 0, rs帐户("帐户支出累计"))
        cur累计进入统筹 = IIf(IsNull(rs帐户("进入统筹累计")), 0, rs帐户("进入统筹累计"))
        cur累计统筹报销 = IIf(IsNull(rs帐户("统筹报销累计")), 0, rs帐户("统筹报销累计"))
    End If
    
    '将此处的数据保存与主程序的数据保存想成一个事务
    '因此就不需要单独的事务控制
    gstrSQL = "zl_帐户年度信息_insert(" & rsTemp("病人ID") & "," & TYPE_重庆松藻 & "," & rsTemp("年度") & "," & _
        cur帐户增加 & "," & cur帐户支出 - rsTemp("个人帐户支付") & "," & cur累计进入统筹 & "," & _
        0 & "," & lng住院次数 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
    
    '冲销单据，处理了几个累计
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_重庆松藻 & "," & rsTemp("病人ID") & "," & _
        rsTemp("年度") & "," & cur帐户增加 & "," & cur帐户支出 - rsTemp("个人帐户支付") & "," & rsTemp("累计进入统筹") * -1 & "," & _
        cur累计统筹报销 & "," & lng住院次数 & "," & rsTemp("起付线") * -1 & "," & rsTemp("封顶线") & "," & rsTemp("实际起付线") * -1 & "," & _
        rsTemp("发生费用金额") * -1 & "," & rsTemp("全自付金额") * -1 & "," & rsTemp("首先自付金额") * -1 & "," & rsTemp("进入统筹金额") * -1 & "," & _
        rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") * -1 & "," & rsTemp("个人帐户支付") * -1 & ",''," & _
        IIf(IsNull(rsTemp("主页ID")), "null", rsTemp("主页ID")) & "," & rsTemp("中途结帐") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
    
    gstrSQL = "select 档次,进入统筹金额,统筹报销金额,比例 from 保险结算计算 where 结帐ID=" & lng结帐ID
    Call OpenRecordset(rs结算计算, "松藻医保")
    
    Do Until rs结算计算.EOF
        '依次为档次、进入统筹金额、统筹报销金额、比例
        gstrSQL = "zl_保险结算计算_Insert(" & lng冲销ID & "," & _
            rs结算计算("档次") & "," & rs结算计算("进入统筹金额") * -1 & "," & rs结算计算("统筹报销金额") * -1 & "," & rs结算计算("比例") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "松藻医保")
        
        rs结算计算.MoveNext
    Loop
    
    住院结算冲销_松藻 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 错误信息_松藻(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息

End Function

Public Function BuildPatiInfo_松藻(ByVal bytType As Byte, ByVal strInfo As String, ByVal lng病人ID As Long) As Long
'功能：建立病人帐户信息
'参数：bytType=0-门诊,1-住院
'      strInfo='0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
'      8中心;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(1,2,3);15退休证号;16年龄段;17灰度级
'      18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型
'返回：病人ID
    Dim rsPati As ADODB.Recordset, str单位编码 As String, lng年龄 As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, curDate As Date
    Dim lng中心 As Long, array信息 As Variant
    
    On Error GoTo errHandle
    If Len(Trim(strInfo)) <> 0 Then
        curDate = zlDatabase.Currentdate
        array信息 = Split(strInfo, ";")
        '从第7项内容中取出单位编码
        If array信息(7) Like "*(*" Then
            str单位编码 = Split(array信息(7), "(")(UBound(Split(array信息(7), "(")))
            str单位编码 = Mid(str单位编码, 1, Len(str单位编码) - 1)
        End If
        '取年龄
        If IsDate(array信息(5)) Then
            lng年龄 = Int(curDate - CDate(array信息(5))) / 365
        End If
        
        lng中心 = Val(array信息(8))
        #If gverControl < 6 Then
            '帐户唯一：险类,中心,医保号
            strSQL = "Select A.*,B.医保号 From 病人信息 A," & _
                " (Select * From 保险帐户" & _
                " Where 险类=" & TYPE_重庆松藻 & _
                " And 医保号='" & CStr(array信息(1)) & "'" & _
                " And 中心=" & lng中心 & ") B" & _
                " Where " & IIf(lng病人ID = 0, "A.病人ID=B.病人ID", "A.病人ID=B.病人ID(+) and A.病人ID=" & lng病人ID) '可能病人ID已经确定
        #Else
            '帐户唯一：险类,中心,医保号
            strSQL = "Select A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.卡验证码, A.费别, A.医疗付款方式, A.姓名, A.性别, A.年龄, A.出生日期, A.出生地点, A.身份证号, A.其他证件, A.身份, A.职业, A.民族, A.国籍, A.区域, A.学历, A.婚姻状况, A.家庭地址," & vbNewLine & _
                "      A.家庭电话, A.家庭地址邮编 As 户口邮编, A.监护人, A.联系人姓名, A.联系人关系, A.联系人地址, A.联系人电话, A.合同单位id, A.工作单位, A.单位电话, A.单位邮编, A.单位开户行, A.单位帐号, A.担保人, A.担保额, A.担保性质, A.就诊时间, A.就诊状态," & vbNewLine & _
                "      A.就诊诊室, A.住院次数, A.当前科室id, A.当前病区id, A.当前床号, A.入院时间, A.出院时间, A.在院, A.Ic卡号, A.健康号, A.医保号, A.险类, A.查询密码, A.登记时间, A.停用时间, A.锁定," & vbNewLine & _
                "      B.医保号 From 病人信息 A," & _
                " (Select * From 保险帐户" & _
                " Where 险类=" & TYPE_重庆松藻 & _
                " And 医保号='" & CStr(array信息(1)) & "'" & _
                " And 中心=" & lng中心 & ") B" & _
                " Where " & IIf(lng病人ID = 0, "A.病人ID=B.病人ID", "A.病人ID=B.病人ID(+) and A.病人ID=" & lng病人ID) '可能病人ID已经确定
        #End If
        Set rsPati = New ADODB.Recordset
        rsPati.CursorLocation = adUseClient
        Call OpenRecordset(rsPati, "松藻医保", strSQL)
        
        If rsPati.EOF Then
            '无保险帐户则认为没有病人信息
            If lng病人ID = 0 Then lng病人ID = GetNextNO(1)
            strSQL = "zl_病人信息_Insert(" & lng病人ID & ",NULL,NULL,NULL," & _
                "'" & array信息(3) & "','" & array信息(4) & "'," & IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
                "To_Date('" & Format(array信息(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "NULL,'" & array信息(6) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
                "NULL,NULL,NULL,NULL,NULL,NULL,'" & array信息(7) & "',NULL,NULL,NULL," & _
                "NULL,NULL,NULL," & TYPE_重庆松藻 & "," & _
                "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call SQLTest(App.ProductName, "松藻医保", strSQL)
            gcnOracle.Execute strSQL, , adCmdStoredProc
            Call SQLTest
        Else
            '有病人信息和保险帐户信息
            If lng病人ID = 0 Then lng病人ID = rsPati!病人ID
            strSQL = "zl_病人信息_Update(" & _
                lng病人ID & "," & IIf(IsNull(rsPati!门诊号), "NULL", rsPati!门诊号) & "," & _
                IIf(IsNull(rsPati!住院号), "NULL", rsPati!住院号) & ",'" & IIf(IsNull(rsPati!费别), "", rsPati!费别) & "'," & _
                "'" & IIf(IsNull(rsPati!医疗付款方式), "", rsPati!医疗付款方式) & "'," & _
                "'" & array信息(3) & "','" & array信息(4) & "'," & IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
                "To_Date('" & Format(array信息(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "'" & IIf(IsNull(rsPati!出生地点), "", rsPati!出生地点) & "','" & array信息(6) & "'," & _
                "'" & IIf(IsNull(rsPati!身份), "", rsPati!身份) & "','" & IIf(IsNull(rsPati!职业), "", rsPati!职业) & "'," & _
                "'" & IIf(IsNull(rsPati!民族), "", rsPati!民族) & "','" & IIf(IsNull(rsPati!国籍), "", rsPati!国籍) & "'," & _
                "'" & IIf(IsNull(rsPati!学历), "", rsPati!学历) & "','" & IIf(IsNull(rsPati!婚姻状况), "", rsPati!婚姻状况) & "'," & _
                "'" & IIf(IsNull(rsPati!家庭地址), "", rsPati!家庭地址) & "','" & IIf(IsNull(rsPati!家庭电话), "", rsPati!家庭电话) & "'," & _
                "'" & IIf(IsNull(rsPati!户口邮编), "", rsPati!户口邮编) & "','" & IIf(IsNull(rsPati!联系人姓名), "", rsPati!联系人姓名) & "'," & _
                "'" & IIf(IsNull(rsPati!联系人关系), "", rsPati!联系人关系) & "','" & IIf(IsNull(rsPati!联系人地址), "", rsPati!联系人地址) & "'," & _
                "'" & IIf(IsNull(rsPati!联系人电话), "", rsPati!联系人电话) & "'," & IIf(IsNull(rsPati!合同单位ID), "NULL", rsPati!合同单位ID) & "," & _
                "'" & array信息(7) & "','" & IIf(IsNull(rsPati!单位电话), "", rsPati!单位电话) & "'," & _
                "'" & IIf(IsNull(rsPati!单位邮编), "", rsPati!单位邮编) & "','" & IIf(IsNull(rsPati!单位开户行), "", rsPati!单位开户行) & "'," & _
                "'" & IIf(IsNull(rsPati!单位帐号), "", rsPati!单位帐号) & "','" & IIf(IsNull(rsPati!担保人), "", rsPati!担保人) & "'," & _
                "" & IIf(IsNull(rsPati!担保额), "NULL", rsPati!担保额) & "," & TYPE_重庆松藻 & ")"
            Call SQLTest(App.ProductName, "松藻医保", strSQL)
            gcnOracle.Execute strSQL, , adCmdStoredProc
            Call SQLTest
        End If
        
        '插入或更新保险帐户信息(自动)
        strSQL = "zl_保险帐户_insert(" & lng病人ID & "," & TYPE_重庆松藻 & "," & _
            lng中心 & "," & _
            "'" & IIf(array信息(0) = "-1", array信息(1), array信息(0)) & "'," & _
            "'" & array信息(1) & "'," & _
            "'" & array信息(2) & "'," & _
            "'" & array信息(9) & "'," & _
            "'" & array信息(15) & "'," & _
            "'" & array信息(10) & "'," & _
            "'" & str单位编码 & "'," & _
            Val(array信息(11)) & "," & _
            Val(array信息(12)) & "," & _
            IIf(Val(array信息(13)) = 0, "NULL", Val(array信息(13))) & "," & _
            IIf(Val(array信息(14)) = 0, 1, Val(array信息(14))) & "," & _
            IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
            "'" & array信息(17) & "'," & _
            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call SQLTest(App.ProductName, "松藻医保", strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
        
        '插入或更新帐户年度信息(自动)
        strSQL = "zl_帐户年度信息_Insert(" & lng病人ID & "," & TYPE_重庆松藻 & "," & Year(curDate) & "," & _
            Val(array信息(18)) & "," & Val(array信息(19)) & "," & _
            Val(array信息(20)) & "," & 0 & "," & Val(array信息(21)) & ")"
        Call SQLTest(App.ProductName, "松藻医保", strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
    End If
    BuildPatiInfo_松藻 = lng病人ID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function 特种病结算(ByVal str疾病名称 As String, ByVal lng在职 As Long, ByVal dbl特准项目 As Currency, Optional ByVal bln住院 As Boolean = False) As String
    Dim dbl个人帐户 As Currency
    Dim dbl基金支付 As Currency, dbl帐户支付 As Currency, dbl单位支付 As Currency
    '返回的串格式与住院虚拟结算一致
    '本函数要修改进入统筹金额及统筹报销金额
    
    dbl个人帐户 = 个人余额_松藻(g结算数据.病人ID)
    dbl个人帐户 = IIf(dbl个人帐户 > 0, dbl个人帐户, 0)
    dbl帐户支付 = 0: dbl基金支付 = 0: dbl单位支付 = 0
    
    Select Case str疾病名称
    Case "精神病", "传染病"
        '（住院）在职、退休，扣除个人帐户，余额全免
        '（门诊）在职、退休，扣除个人帐户，仅免特准项目，除特准项目外的，仍按医保规则结算
        If bln住院 Then
            If dbl个人帐户 > 0 Then
                If g结算数据.进入统筹金额 <= dbl个人帐户 Then
                    dbl帐户支付 = g结算数据.进入统筹金额
                Else
                    dbl帐户支付 = dbl个人帐户
                End If
            Else
                dbl帐户支付 = dbl个人帐户
            End If
            dbl基金支付 = g结算数据.进入统筹金额 - dbl帐户支付
        End If
    Case "职业病", "癌症"
        '（住院、门诊）扣除个人帐户，余额全免
        If dbl个人帐户 > 0 Then
            If g结算数据.进入统筹金额 <= dbl个人帐户 Then
                dbl帐户支付 = g结算数据.进入统筹金额
            Else
                dbl帐户支付 = dbl个人帐户
            End If
        Else
            dbl帐户支付 = dbl个人帐户
        End If
        dbl基金支付 = g结算数据.进入统筹金额 - dbl帐户支付
    Case "工伤", "计划生育"
        '（住院、门诊）全免，不扣除个人帐户
        dbl基金支付 = g结算数据.进入统筹金额
    Case "单位支付"
        dbl单位支付 = g结算数据.发生费用金额
    Case "家属病人"
        '（住院），药费、手术费、血费由医保基金支付50%，余下的自费
        '（门诊），药费由医保基金支付50%，余下的自费
    End Select
    
    dbl基金支付 = Val(Format(dbl基金支付, "#####0.00;-#####0.00;0;"))
    dbl帐户支付 = Val(Format(dbl帐户支付, "#####0.00;-#####0.00;0;"))
    dbl单位支付 = Val(Format(dbl单位支付, "#####0.00;-#####0.00;0;"))
    g结算数据.统筹报销金额 = dbl基金支付
    g结算数据.个人帐户支付 = dbl帐户支付
    
    '组合返回串
    特种病结算 = "医保基金;" & g结算数据.统筹报销金额 & ";0|个人帐户;" & g结算数据.个人帐户支付 & ";0"
    If dbl单位支付 <> 0 Then 特种病结算 = 特种病结算 & "|单位支付;" & dbl单位支付 & ";0"
End Function


