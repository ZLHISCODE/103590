Attribute VB_Name = "mdl中梁山"
Option Explicit

Public Function 医保初始化_中梁山() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false

    
    '为了避免授权难度增加，此处不再进行对各个医保表数据的检查
    医保初始化_中梁山 = True
End Function

Public Function 身份标识_中梁山(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    
    Dim strTmpIden As String
    
    strTmpIden = frmIdentify中梁山.ShowCard(bytType, lng病人ID)
    身份标识_中梁山 = strTmpIden
End Function

Public Function 个人余额_中梁山(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'参数: bytYear-余额类型,0-所有余额,1-本年余额,2-往年余额
'返回: 返回个人帐户余额的金额
    
    '不使用个人帐户
    个人余额_中梁山 = 0
End Function

Public Function 门诊结算_中梁山(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, _
            ByVal cur全自费 As Currency, ByVal cur首先自付 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
    
    门诊结算_中梁山 = False
End Function


Public Function 门诊结算冲销_中梁山(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额

    门诊结算冲销_中梁山 = False
End Function

Public Function 个人帐户转预交_中梁山(lng预交ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    
    个人帐户转预交_中梁山 = False
End Function


Public Function 个人帐户转预交冲销_中梁山(lng预交ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false

    个人帐户转预交冲销_中梁山 = False
End Function

Public Function 入院登记_中梁山(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false

    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆中梁山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    入院登记_中梁山 = True
End Function

Public Function 出院登记_中梁山(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重庆中梁山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    出院登记_中梁山 = True
End Function

Public Function 出院登记撤消_中梁山(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
            '取入院登记验证所返回的顺序号
    On Error GoTo errHandle
    
        
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆中梁山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    出院登记撤消_中梁山 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_中梁山(rs费用明细 As Recordset, ByVal lng病人ID As Long) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '结算要求：NO、序号、病人ID、医保项目编码、收费类别、收费名称、开单部门、规格、产地、数量、价格、金额、医生,登记时间(发生时间),婴儿费,保险大类ID
    Dim rs特准项目 As New ADODB.Recordset
    Dim rs算法 As New ADODB.Recordset          '保存
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng中心 As Long
    Dim lng在职 As Long, lng年龄段 As Long, lng年龄 As Long
    
    Dim dbl最大金额  As Double ''对一个按住院日计算的项目，最多能得到的金额
    
    Dim cur全自费 As Currency, cur首先自付 As Currency, cur进入统筹 As Currency, dblTemp As Double
    Dim bln全额统筹 As Boolean, bln无封顶线 As Boolean, bln目录 As Boolean, bln存在诊疗目录 As Boolean
    
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '1、初始化一些变量
    With g结算数据
        .超限自付金额 = 0
        .发生费用金额 = 0
        .封顶线 = 0
        .进入统筹金额 = 0
        .累计进入统筹 = 0
        .累计统筹报销 = 0
        .全自费金额 = 0
        .首先自付金额 = 0
        .统筹报销金额 = 0
        .实际起付线 = 0    '用于保存进入全报目录的金额
        
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
    
    '1.2 读出病人的入院时间
    gstrSQL = "select 入院日期,nvl(出院日期,to_date('3000-01-01','yyyy-MM-dd')) as 出院日期 " & _
              "from 病案主页 where 病人ID=" & g结算数据.病人ID & " and 主页ID=" & g结算数据.主页ID
    Call OpenRecordset(rsTemp, "虚拟结算")
    If rsTemp("出院日期") = CDate("3000-01-01") Then
        g结算数据.中途结帐 = 1
    Else
        '表示该病人已经出院
        g结算数据.中途结帐 = 0
    End If

    '1.3 读出本次住院期间累计结帐情况
    With g结算数据
        gstrSQL = "select A.中心,A.人员身份,A.在职,A.年龄段," & _
                  "      B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计" & _
                  " from 保险帐户 A,帐户年度信息 B" & _
                  " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) " & _
                  "     and B.年度(+)=" & .年度 & " and A.病人ID=" & .病人ID & " and A.险类=" & TYPE_重庆中梁山
        Call OpenRecordset(rsTemp, "虚拟结算")
        
        lng中心 = IIf(IsNull(rsTemp("中心")), 0, rsTemp("中心"))
        lng在职 = IIf(IsNull(rsTemp("在职")), 1, rsTemp("在职"))
        lng年龄 = IIf(IsNull(rsTemp("年龄段")), 0, rsTemp("年龄段"))
        .住院次数 = IIf(IsNull(rsTemp("住院次数累计")), 0, rsTemp("住院次数累计"))
        .累计进入统筹 = IIf(IsNull(rsTemp("进入统筹累计")), 0, rsTemp("进入统筹累计"))
        .累计统筹报销 = IIf(IsNull(rsTemp("统筹报销累计")), 0, rsTemp("统筹报销累计"))
    
        gstrSQL = "select 年龄段,nvl(全额统筹,0) as 全额统筹 ,nvl(无起付线,0) as 无起付线 ,nvl(无封顶线,0) as 无封顶线 " & _
                " from 保险年龄段" & _
                " where 险类=" & TYPE_重庆中梁山 & " and nvl(中心,0)=" & lng中心 & _
                "       and 在职=" & lng在职 & " and 下限<=" & lng年龄 & " and (" & lng年龄 & "<=上限 or 上限=0)"
        Call OpenRecordset(rsTemp, "虚拟结算")
        If rsTemp.RecordCount = 0 Then
            MsgBox "请在“保险类别管理”中设置年龄段与费用档。", vbInformation, gstrSysName
            Exit Function
        End If
        lng年龄段 = rsTemp("年龄段")
        bln全额统筹 = (rsTemp("全额统筹") = 1)
        bln无封顶线 = (rsTemp("无封顶线") = 1)
    End With
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '2、按统筹支付项目合计发生金额和数量
    '2.1、根据病人病种，判断是否有全保的项目
    '2.2、计算进入统筹金额
    If bln全额统筹 = False Then
        gstrSQL = "SELECT B.收费细目ID " & _
                 " FROM 保险帐户 A,保险特准项目 B,收费细目 C " & _
                 " WHERE A.病人ID=" & g结算数据.病人ID & " AND A.险类=" & TYPE_重庆中梁山 & " AND A.病种ID=B.病种ID And B.收费细目ID=C.ID and C.类别 not in ('5','6','7') and Rownum<2"
        Call OpenRecordset(rsTemp, "住院结算")
        If rsTemp.RecordCount > 0 Then bln存在诊疗目录 = True '如果选择有诊疗项目，才认为特准项目清单有效
        
        gstrSQL = "SELECT B.收费细目ID " & _
                 " FROM 保险帐户 A,保险特准项目 B " & _
                 " WHERE A.病人ID=" & g结算数据.病人ID & " AND A.险类=" & TYPE_重庆中梁山 & " AND A.病种ID=B.病种ID"
        Call OpenRecordset(rs特准项目, "住院结算")
        
        gstrSQL = "select ID,算法,统筹比额,特准定额,特准天数,是否医保 FROM 保险支付大类  where 险类=" & TYPE_重庆中梁山
        Call OpenRecordset(rs算法, "住院结算")
        
        dblTemp = 0
        If rs费用明细.RecordCount > 0 Then rs费用明细.MoveFirst
        Do Until rs费用明细.EOF
            bln目录 = False
            rs特准项目.Filter = "收费细目ID=" & rs费用明细("收费细目ID")
            
            If rs特准项目.RecordCount = 0 Then
                '没有设置特定项目
                If rs费用明细("收费类别") = "5" Or rs费用明细("收费类别") = "6" Or rs费用明细("收费类别") = "7" Then
                    '药品
                    bln目录 = False
                ElseIf bln存在诊疗目录 = True Then
                    '诊疗
                    bln目录 = True
                End If
            Else
                If rs费用明细("收费类别") = "5" Or rs费用明细("收费类别") = "6" Or rs费用明细("收费类别") = "7" Then
                    '药品
                    bln目录 = True
                Else
                    '诊疗
                    bln目录 = False
                End If
            End If
            
            If bln目录 = False Then
                '不在特准项目中，只有按比例进行计算
                rs算法.Filter = "ID=" & rs费用明细("保险大类ID")
                If rs算法.RecordCount > 0 Then
                    '算法:1-总额计算项目；2-住院日核定项目
                    If rs算法("算法") = 1 Then
                        If rs算法("统筹比额") = 0 Then
                            cur全自费 = cur全自费 + rs费用明细("金额")
                        Else
                            cur进入统筹 = cur进入统筹 + rs费用明细("金额") * rs算法("统筹比额") / 100
                        End If
                    Else
                        If Val(rs费用明细("数量")) > Val(rs算法("特准天数")) Then
                            '如果住院日超过特准天数，那么最大金额就是 特准天数*特准定额 +  (数量-特准天数)*统筹比额
                            '当特准定额或特准天数任一个为0时，就相当于不要特准天数
                            dbl最大金额 = rs算法("特准定额") * rs算法("特准天数") + _
                                (rs费用明细("数量") - IIf(rs算法("特准定额") = 0 Or rs算法("特准天数") = 0, 0, rs算法("特准天数"))) * rs算法("统筹比额")
                        Else
                            '如果住院日低于特准天数，那么最大金额就是 数量*特准定额 或者 数量*统筹比额
                            '当特准定额或特准天数任一个为0时，就相当于不要特准定额
                            If rs算法("特准定额") = 0 Or rs算法("特准天数") = 0 Then
                                dbl最大金额 = rs费用明细("数量") * rs算法("统筹比额")
                            Else
                                dbl最大金额 = rs费用明细("数量") * rs算法("特准定额")
                            End If
                        End If
                        
                        '总金额比最大金额小，就取全部金额；否则只最大金额
                        cur进入统筹 = cur进入统筹 + IIf(rs费用明细("金额") < dbl最大金额, rs费用明细("金额"), dbl最大金额)
                        
                        If rs费用明细("金额") > dbl最大金额 Then
                            '全部算作全自费
                            cur全自费 = cur全自费 + rs费用明细("金额") - dbl最大金额
                        End If
                    End If
                Else
                    cur全自费 = cur全自费 + rs费用明细("金额")
                End If
            Else
                cur进入统筹 = cur进入统筹 + rs费用明细("金额")
                g结算数据.实际起付线 = g结算数据.实际起付线 + rs费用明细("金额")
            End If
            
            dblTemp = dblTemp + rs费用明细("金额")
            rs费用明细.MoveNext
        Loop
        
        g结算数据.发生费用金额 = dblTemp
        g结算数据.进入统筹金额 = cur进入统筹
        g结算数据.全自费金额 = cur全自费
        g结算数据.首先自付金额 = g结算数据.发生费用金额 - cur全自费 - cur进入统筹
    Else
        Do Until rs费用明细.EOF
                
            dblTemp = dblTemp + rs费用明细("金额")
            rs费用明细.MoveNext
        Loop
        g结算数据.发生费用金额 = dblTemp
        g结算数据.进入统筹金额 = g结算数据.发生费用金额
        g结算数据.全自费金额 = 0
        g结算数据.首先自付金额 = 0
    End If
        
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '3、获得起付线、封顶线、支付比例等数据
    '3.1、获得起付线、封顶线
    With g结算数据
        If bln无封顶线 = True Then
            .封顶线 = 0
        Else
            '检查病人的病种是否有特殊封顶线
            gstrSQL = "SELECT B.特殊封顶线,b.封顶线金额 " & _
                     " FROM 保险帐户 A,保险病种 B " & _
                     " WHERE A.病人ID=" & g结算数据.病人ID & " AND A.险类=" & TYPE_重庆中梁山 & " AND A.病种ID=B.ID(+)"
            Call OpenRecordset(rsTemp, "虚拟结算")
            
            If Nvl(rsTemp("特殊封顶线"), 0) = 1 Then
                If IsNull(rsTemp("封顶线金额")) = True Then
                    bln无封顶线 = True '这种情况也属于无封顶线，如工伤、矽肺
                Else
                    .封顶线 = rsTemp("封顶线金额")
                End If
            Else
                gstrSQL = "select max(decode(A.性质,'A',A.金额,0)) as 封项线 " & _
                          "  from 保险支付限额 A " & _
                          "  where A.险类=" & TYPE_重庆中梁山 & " and A.中心=" & lng中心 & " and A.性质='A' and A.年度=" & .年度
                Call OpenRecordset(rsTemp, "虚拟结算")
                        
                .封顶线 = IIf(IsNull(rsTemp("封项线")), 0, rsTemp("封项线"))
                If .封顶线 = 0 Then
                    MsgBox "请在“年度结算规则”中设置本年度的封顶线。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End With
    
    '3.3、取得费用档次
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "select B.档次,B.下限,B.上限,A.比例 " & _
              "  from 保险支付比例 A,保险费用档 B " & _
              "  Where A.险类 =" & TYPE_重庆中梁山 & " And A.中心 =" & lng中心 & " And A.年度 =" & g结算数据.年度 & " And A.在职 =" & lng在职 & " And A.年龄段 =" & lng年龄段 & _
              "       and A.险类=B.险类 and A.中心=b.中心 and A.档次=B.档次 and B.档次=1"
    Call OpenRecordset(rsTemp, "虚拟结算")
    If rsTemp.RecordCount = 0 Then
        MsgBox "请在“年度结算规则”中设置本年度的统筹支付比例。。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '4、计算该次结算可报销的金额
    With g结算数据
        If bln无封顶线 = True Then
            '不用考虑有报不了的金额
            .统筹报销金额 = .实际起付线 + (.进入统筹金额 - .实际起付线) * rsTemp("比例") / 100
        Else
            '计算出可能报销的金额。进入特定目录的，不计算比较
            dblTemp = .实际起付线 + (.进入统筹金额 - .实际起付线) * rsTemp("比例") / 100
            If dblTemp > .封顶线 - .累计统筹报销 Then
                .统筹报销金额 = .封顶线 - .累计统筹报销
                .超限自付金额 = .进入统筹金额 - .统筹报销金额 / rsTemp("比例") * 100
                If .超限自付金额 < 0 Then .超限自付金额 = 0                   '如果进入统筹金额还不到起付线，为负数
            Else
                .统筹报销金额 = dblTemp
            End If
        End If
    End With
    
    住院虚拟结算_中梁山 = "医保基金;" & g结算数据.统筹报销金额 & ";0"
End Function

Public Function 住院结算_中梁山(lng结帐ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    Dim rsTemp As New ADODB.Recordset
    Dim var结算计算 As Variant
        
    With g结算数据
        gstrSQL = "zl_帐户年度信息_insert(" & .病人ID & "," & TYPE_重庆中梁山 & "," & .年度 & "," & _
            .帐户累计增加 & "," & .帐户累计支出 & "," & .累计进入统筹 + .进入统筹金额 & "," & _
            .累计统筹报销 + .统筹报销金额 & "," & .住院次数 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
        
        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_重庆中梁山 & "," & .病人ID & "," & _
            .年度 & "," & .帐户累计增加 & "," & .帐户累计支出 & "," & .累计进入统筹 & "," & _
            .累计统筹报销 & "," & .住院次数 + 1 & "," & .起付线 & "," & .封顶线 & "," & .实际起付线 & "," & _
            .发生费用金额 & "," & .全自费金额 & "," & .首先自付金额 & "," & .进入统筹金额 & "," & .统筹报销金额 & ",0," & _
            .超限自付金额 & ",0,NULL," & .主页ID & "," & .中途结帐 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    End With
    
    住院结算_中梁山 = True
End Function

Public Function 住院结算冲销_中梁山(lng结帐ID As Long) As Boolean
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
    Call OpenRecordset(rsTemp, "模拟医保")
    
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "Select * " & _
              "  From 保险结算记录 Where 性质=2 and 记录ID='" & lng结帐ID & "'"
    Call OpenRecordset(rsTemp, "模拟医保")
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "该病人的医保结算数据丢失，不能作废。"
        Exit Function
    End If
    If Can住院结算冲销(rsTemp("病人ID"), rsTemp("主页ID")) = False Then Exit Function
    
    gstrSQL = "select B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计 " & _
              " from 保险帐户 A,帐户年度信息 B " & _
              " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) and B.年度(+)=" & Year(zlDatabase.Currentdate) & " and A.病人ID=" & rsTemp("病人ID") & " and A.险类=" & TYPE_重庆中梁山
    Call OpenRecordset(rs帐户, "模拟医保")
    
    If rs帐户.EOF = False Then
        lng住院次数 = IIf(IsNull(rs帐户("住院次数累计")), 0, rs帐户("住院次数累计"))
        cur帐户增加 = IIf(IsNull(rs帐户("帐户增加累计")), 0, rs帐户("帐户增加累计"))
        cur帐户支出 = IIf(IsNull(rs帐户("帐户支出累计")), 0, rs帐户("帐户支出累计"))
        cur累计进入统筹 = IIf(IsNull(rs帐户("进入统筹累计")), 0, rs帐户("进入统筹累计"))
        cur累计统筹报销 = IIf(IsNull(rs帐户("统筹报销累计")), 0, rs帐户("统筹报销累计"))
    End If
    
    '将此处的数据保存与主程序的数据保存想成一个事务
    '因此就不需要单独的事务控制
    gstrSQL = "zl_帐户年度信息_insert(" & rsTemp("病人ID") & "," & TYPE_重庆中梁山 & "," & rsTemp("年度") & "," & _
        cur帐户增加 & "," & cur帐户支出 - rsTemp("个人帐户支付") & "," & cur累计进入统筹 - rsTemp("进入统筹金额") & "," & _
        cur累计统筹报销 - rsTemp("统筹报销金额") & "," & lng住院次数 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    '冲销单据，处理了几个累计
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_重庆中梁山 & "," & rsTemp("病人ID") & "," & _
        rsTemp("年度") & "," & cur帐户增加 & "," & cur帐户支出 - rsTemp("个人帐户支付") & "," & cur累计进入统筹 - rsTemp("进入统筹金额") & "," & _
        cur累计统筹报销 - rsTemp("统筹报销金额") & "," & lng住院次数 & "," & rsTemp("起付线") * -1 & "," & rsTemp("封顶线") & "," & rsTemp("实际起付线") * -1 & "," & _
        rsTemp("发生费用金额") * -1 & "," & rsTemp("全自付金额") * -1 & "," & rsTemp("首先自付金额") * -1 & "," & rsTemp("进入统筹金额") * -1 & "," & _
        rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") * -1 & "," & rsTemp("个人帐户支付") * -1 & ",''," & _
        IIf(IsNull(rsTemp("主页ID")), "null", rsTemp("主页ID")) & "," & rsTemp("中途结帐") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
    
    
    gstrSQL = "select 档次,进入统筹金额,统筹报销金额,比例 from 保险结算计算 where 结帐ID=" & lng结帐ID
    Call OpenRecordset(rs结算计算, "模拟医保")
    
    Do Until rs结算计算.EOF
        '依次为档次、进入统筹金额、统筹报销金额、比例
        gstrSQL = "zl_保险结算计算_Insert(" & lng冲销ID & "," & _
            rs结算计算("档次") & "," & rs结算计算("进入统筹金额") * -1 & "," & rs结算计算("统筹报销金额") * -1 & "," & rs结算计算("比例") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "模拟医保")
        
        rs结算计算.MoveNext
    Loop
    
    住院结算冲销_中梁山 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 记帐传输_中梁山(strNO As String, int性质 As Integer, int状态 As Integer, Optional lng病人ID As Long) As Boolean
'功能：将住院病人的记帐单据上传到医保前置服务器
'参数：lng病人ID=是否只上传单据中指定病人的费用
    
    '暂时不作任何处理。
    '如果用户要求打印出现的病人一日清单要体现特殊病诊疗目录中的统筹比例，那在此处理修改
    记帐传输_中梁山 = True
End Function




