Attribute VB_Name = "mdlBill"
Option Explicit

'本模块是专门为收费记帐而建立的
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public gbln简码切换 As Boolean
Public gcolPrivs As Collection              '记录内部模块的权限
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    部门名称 As String
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

'内部应用模块号定义
Public Enum Enum_Inside_Program
    p住院记帐 = 1133
    p病人结帐 = 1137
    p费用查询 = 1139
    p一日清单 = 1141
    p记帐操作 = 1150
End Enum

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
    support负数记帐 = 35            '是否允许负数记帐，操作员首先要拥有负数记帐的权限。此参数缺省为真，不支持的接口需单独处理
    
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    
    support实时监控 = 60             '是否启用费用实时监控
End Enum

'系统公用临时变量
'Public glngOK As Long   '如果成功就返回1；取消就返回0；如果是由于记帐单模板被删除，返回-1
Public gblnOK As Boolean
Public gbytWarn As Byte '记帐报警返回值
Public gstrModiNO As String '修改后产生的新单据号

'============医保参数=====================
Public gclsInsure As New clsInsure
'============费用系统参数=====================
Public grsPar As ADODB.Recordset '记录系统参数

'刷卡控制
Public gbytCardNOLen As Byte '就诊卡号长度
Public gblnShowCard As Boolean '是否就诊卡号显示为正常符号
Public gstrCardPass As String '刷卡时要求输入密码,'0000000000'依位顺序表示各个场合,分别为:1.门诊挂号,2.门诊划价,3.门诊收费,4.门诊记帐,5.入院登记,6.住院记帐,7.病人结帐,8.病人预交款,9.检验技师站,10.影像医技站.'

'单据输入控制
Public gstrLike As String  '项目匹配方法,%或空
Public gbytCode As Byte '简码生成方式，0-拼音,1-五笔,2-两者
Public gblnMyStyle As Boolean '使用个性化风格
Public gstrIme As String '自动的开启输入法

Public gstrMatchMode As String '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
Public gbyt医保对码检查 As Byte '0-不进行检查、1-检查并提醒未对码项目、2-检查并禁止未对码项目
Public gcurMaxMoney As Currency '单笔费用最大提醒金额

'操作控制
Public gbytBillOpt As Byte '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。

'其它控制
Public gblnDailyTime As Boolean '真时表示允许时间超过服务器时间
Public gbln报警包含划价费用 As Boolean '记帐报警包含划价费用

Public gintOutDay As Integer '结帐可选择出院病人天数

'输入控制
Public gstr收费类别 As String '可输入的收费类别
Public gblnTime As Boolean '变价是否可以输入数次
Public gbln护士 As Boolean '开单人是否显示护士
Public gbln开单人 As Boolean '记帐是否必须输入开单人

'留观病人记帐
Public gbln门诊留观 As Boolean
Public gbln住院留观 As Boolean
'刘兴洪 问题:????    日期:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '费用小数精度
Public gstrFeePrecisionFmt As String '费用小数格式:0.00000

'金额小数位数
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"

'系统参数:34717
Public Type TY_Reg_Para  '挂号相关参数
    bytNODaysGeneral As Byte    '普通挂号有效天数
    bytNoDayseMergency As Byte '急诊挂号有效天数
End Type
Public Type TY_SysPara
    Sy_Reg  As TY_Reg_Para
    byt病人审核方式 As Byte '49501:病人审核方式:0-未审核不允许结帐，缺省为0;1-审核时不许调整费用和医嘱（包含医嘱调整和费用调整）
    bln未入科禁止记账 As Boolean '51612
End Type
Public gSysPara As TY_SysPara       '系统参数相关;以后可以扩展(刘兴洪)

Private mlng部门编码平均长度 As Long

Public Function BillingWarn(strPrivs As String, str姓名 As String, lng病区ID As Long, str适用病人 As String, _
    rsWarn As ADODB.Recordset, cur余额 As Currency, cur当日额 As Currency, _
    cur单据金额 As Currency, cur担保 As Currency, str类别 As String, _
    ByVal str类别名 As String, ByRef str已报类别 As String, Optional bln多病人 As Boolean, _
    Optional curItemMoney As Currency = 0, _
    Optional blnNotCheck类别 As Boolean = False) As Integer
'功能:对病人记帐进行报警提示
'参数:
'     str姓名=病人姓名,用于提示
'     lng病区ID=病人病区ID,用于选择适用的病区报警设置，0表示没有确定病区，仅所有病区的报警设置适用
'     str适用病人=根据病人身份返回的记帐报警适用方案
'     rsWarn=当前病区记帐报警设置记录
'     cur余额=病人余额,用于累计报警
'     cur当日额=病人当日发生的费用额,用于每日报警
'     cur单据金额=病人单据中输入的费用
'     cur担保=病人担保费用额,用于累计报警
'     str类别=当前要检查的类别,用于分类报警
'     str类别名=类别名称,用于提示
'     curItemMoney-当笔金额(如果传入<>0 ,则需要判断当笔情况,如果超出金额,则允许用户继续,否则根据报警方式进行):刘兴洪:24491
'     blnNotCheck类别:不对类别进行检查(主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
'返回:0;没有报警,继续
'     1:报警提示后用户选择继续
'     2:报警提示后用户选择中断
'     3:报警提示必须中断
'     4:强制记帐报警,继续
'     str报警类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
    Dim i As Integer, byt标志 As Byte
    Dim bln已报警 As Boolean
    Dim byt方式 As Byte, byt已报方式 As Byte
    Dim arrTmp As Variant
    
    On Error GoTo errH
    
    '报警参数检查
    If rsWarn.State = 0 Then Exit Function '20030709
    rsWarn.Filter = "适用病人='" & str适用病人 & "' And 病区ID=" & lng病区ID
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!报警值) Then Exit Function
    
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str类别) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志1 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str类别) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志2 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str类别) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志3 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    
    If byt标志 = 0 Then Exit Function '无有效设置
    
    '报警标志2实际上是两种判断①②,其它只有一种判断①
    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
    If bln多病人 Then
        '示例：",周:-,张:DEF,李:567,张567"
        '报警标志2示例：",周:-①,张:DEF①,李:567①,张567②"
        bln已报警 = str已报类别 & "," Like "*," & str姓名 & ":-*,*" _
            Or str已报类别 & "," Like "*," & str姓名 & ":*" & str类别 & "*,*"
    Else
        '示例："-" 或 ",ABC,567,DEF"
        '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
        bln已报警 = InStr(str已报类别, str类别) > 0 Or str已报类别 Like "-*"
    End If
    
    If bln已报警 Then
        If byt标志 = 2 Then
            If bln多病人 Then
                arrTmp = Split(str已报类别, ",")
                For i = 0 To UBound(arrTmp)
                    If "," & arrTmp(i) & "," Like "*," & str姓名 & ":-*,*" _
                        Or "," & arrTmp(i) & "," Like "*," & str姓名 & ":*" & str类别 & "*,*" Then
                        byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                        Exit For
                    End If
                Next
            Else
                If str已报类别 Like "-*" Then
                    byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
                Else
                    arrTmp = Split(str已报类别, ",")
                    For i = 0 To UBound(arrTmp)
                        If InStr(arrTmp(i), str类别) > 0 Then
                            byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                            Exit For
                        End If
                    Next
                End If
            End If
        Else
            Exit Function
        End If
    End If
    
    If str类别名 <> "" Then str类别名 = """" & str类别名 & """费用"
        
    '---------------------------------------------------------------------
    If rsWarn!报警方法 = 1 Then  '累计费用报警(低于)
        Select Case byt标志
            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                If cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If curItemMoney <> 0 And cur余额 + cur担保 - (cur单据金额 - curItemMoney) > Val(Nvl(rsWarn!报警值)) Then
                         '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        If MsgBox("注意" & vbCrLf & _
                                   "    病人“" & str姓名 & "” 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ", 是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                    End If
                
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1
                        End If
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & " 低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If cur余额 + cur担保 - cur单据金额 < 0 Then
                        '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                         If curItemMoney <> 0 And cur余额 + cur担保 - (cur单据金额 - curItemMoney) > 0 Then
                             '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                             '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                            If MsgBox("注意" & vbCrLf & _
                                       "    病人“" & str姓名 & "” 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽,是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                            End If
                            Exit Function
                         End If
                                             
                        byt方式 = 2
                        If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                            MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。", vbInformation, gstrSysName
                            BillingWarn = 3
                        Else
                            MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
                            BillingWarn = 4
                        End If
                    ElseIf cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                        '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                         If curItemMoney <> 0 And cur余额 + cur担保 - (cur单据金额 - curItemMoney) > Val(Nvl(rsWarn!报警值)) Then
                             '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                             '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                            If MsgBox("注意" & vbCrLf & _
                                       "    病人“" & str姓名 & "” 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ", 是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                            End If
                            Exit Function
                         End If
                    
                    
                        byt方式 = 1
                        If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                            If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1
                            End If
                        Else
                            MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If cur余额 + cur担保 - cur单据金额 < 0 Then
                            '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                             If curItemMoney <> 0 And cur余额 + cur担保 - (cur单据金额 - curItemMoney) > 0 Then
                                '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                                '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                               If MsgBox("注意" & vbCrLf & _
                                          "    病人“" & str姓名 & "” 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽,是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                   BillingWarn = 2
                               Else
                                   BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                               End If
                               Exit Function
                            End If
                         
                            byt方式 = 2
                            If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                                MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。", vbInformation, gstrSysName
                                BillingWarn = 3
                            Else
                                MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt已报方式 = 2 Then
                        '上次预交款已经耗尽并强制继续,不再处理
                        Exit Function
                    End If
                End If
            Case 3 '低于报警值禁止记帐
                If cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If curItemMoney <> 0 And cur余额 + cur担保 - (cur单据金额 - curItemMoney) > Val(Nvl(rsWarn!报警值)) Then
                         '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        If MsgBox("注意" & vbCrLf & _
                                   "    病人“" & str姓名 & "” 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽,是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                    
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                        BillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If cur当日额 + cur单据金额 > rsWarn!报警值 Then
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If curItemMoney <> 0 And cur当日额 + cur单据金额 - curItemMoney < Val(Nvl(rsWarn!报警值)) Then
                         '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        If MsgBox("注意" & vbCrLf & _
                                   "    病人“" & str姓名 & "” 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        If MsgBox(str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, "0.00") & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1
                        End If
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, "0.00") & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If cur当日额 + cur单据金额 > rsWarn!报警值 Then
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If curItemMoney <> 0 And cur当日额 + cur单据金额 - curItemMoney < Val(Nvl(rsWarn!报警值)) Then
                         '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        If MsgBox("注意" & vbCrLf & _
                                   "    病人“" & str姓名 & "” 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        MsgBox str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, "0.00") & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                        BillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, "0.00") & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '对于继续类的操作,返回已报警类别
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt标志 = 1 Then
            If rsWarn!报警标志1 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志1
            End If
        ElseIf byt标志 = 2 Then
            If rsWarn!报警标志2 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志2
            End If
            '附加标注以判断已报警的具体方式
            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
        ElseIf byt标志 = 3 Then
            If rsWarn!报警标志3 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志3
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub GetDoctor(lng科室ID As Long, ByVal bln护士 As Boolean, ByRef rsTmp As ADODB.Recordset, ByVal int病人来源 As Integer)
'功能：获取指定科室的医生
'参数：lng科室ID=指定科室ID,bln护士=是否也读取护士(收费\划价)
    'Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '允许部门工作性质为非临床的医生或护士,因为可能该人所属的一个部门是非末级部门.
    If rsTmp Is Nothing Then
        strSQL = _
            "Select Distinct A.ID,B.部门ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
            " C.人员性质,Nvl(A.聘任技术职务,0) as 职务,B.缺省" & _
            " From 人员表 A,部门人员 B,人员性质说明 C,部门性质说明 D" & _
            " Where A.ID = B.人员ID And A.ID=C.人员ID And B.部门ID=D.部门ID And C.人员性质 IN('医生','护士') " & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " And D.服务对象 IN(" & int病人来源 & ",3) And D.工作性质 IN('临床','手术') And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & _
            " Order by 简码,缺省 Desc"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlBill")
    End If
   
    If lng科室ID = 0 Then
        rsTmp.Filter = IIf(bln护士, "", "人员性质='医生'")
    Else
        rsTmp.Filter = "部门ID=" & lng科室ID & IIf(bln护士, "", " And 人员性质='医生'")
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Public Function MakeDetailRecord(ByRef objBill As ExpenseBill, ByVal str开单人 As String, ByVal str开单科室 As String, _
    Optional ByVal lngRow As Long = -1) As ADODB.Recordset
'功能：根据单据对象内容创建一个明细记录集信息(以售价单位)
'字段：病人ID，主页ID，收费类别，收费细目ID，数量，单价，实收金额，开单人，开单科室
'参数：intPage=指定的单据,lngRow=指定的行，不指定时包含所有单据的所有行,注意行号是从0开始的，并且对象中加了R前缀
    Dim i As Integer, j As Integer
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl单价 As Double, cur实收 As Currency
    Dim rsTmp As New ADODB.Recordset
    '79420,李南春,2014/11/10:调整记录集字段大小
    rsTmp.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "数量", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "单价", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    
    If lngRow = -1 Then
        intB = 0
        intE = objBill.Details.Count - 1
    Else
        intB = lngRow
        intE = lngRow
    End If
    
    For i = intB To intE
        dbl单价 = 0: cur实收 = 0
        With objBill.Details("R" & i)
            If .收费细目ID <> 0 Then    '传入的objBill中的明细可能还没有输信息，是预设计的行数
                If lngRow = -1 Then
                    rsTmp.Filter = "收费细目ID=" & .收费细目ID
                    blnNew = rsTmp.RecordCount = 0
                Else
                    blnNew = True
                End If
                                
                If blnNew Then
                    rsTmp.AddNew
                    
                    rsTmp!病人ID = objBill.病人ID
                    rsTmp!主页ID = objBill.主页ID
                    
                    rsTmp!收费类别 = .收费类别
                    rsTmp!收费细目ID = .收费细目ID
                    
                    
                    For j = 1 To .InComes.Count
                        dbl单价 = dbl单价 + .InComes(j).标准单价
                        cur实收 = cur实收 + .InComes(j).实收金额
                    Next
                    rsTmp!数量 = .数次
                    rsTmp!单价 = Format(dbl单价, gstrFeePrecisionFmt)
                    
                    rsTmp!实收金额 = Format(cur实收, gstrDec)
                    
                    rsTmp!开单人 = str开单人
                    rsTmp!开单科室 = str开单科室
                Else
                    For j = 1 To .InComes.Count
                        dbl单价 = dbl单价 + .InComes(j).标准单价
                        cur实收 = cur实收 + .InComes(j).实收金额
                    Next
                    rsTmp!数量 = rsTmp!数量 + .数次
                    rsTmp!单价 = Format((rsTmp!单价 + Format(dbl单价, gstrFeePrecisionFmt)) / 2, gstrFeePrecisionFmt)
                    
                    rsTmp!实收金额 = rsTmp!实收金额 + Format(cur实收, gstrDec)
                End If
                
                rsTmp.Update
            End If
        End With
    Next
    
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Sub LoadPatientBaby(ByRef cboBaby As ComboBox, ByVal lngPatient As Long, lngPatientPage As Long)
    Dim rsTmp As ADODB.Recordset, i As Long
    
    cboBaby.Clear
    cboBaby.AddItem "0-病人本人"
    cboBaby.ItemData(cboBaby.NewIndex) = 0
    Call zlControl.CboSetIndex(cboBaby.hwnd, 0)
    
    If lngPatient <> 0 Then
        Set rsTmp = GetPatientBaby(lngPatient, lngPatientPage)
        With rsTmp
            For i = 1 To .RecordCount
                If Not IsNull(!婴儿姓名) Then
                    cboBaby.AddItem !序号 & "-" & !婴儿姓名
                Else
                    cboBaby.AddItem !序号 & "-第" & !序号 & "个婴儿"
                End If
                cboBaby.ItemData(cboBaby.NewIndex) = !序号
                .MoveNext
            Next
        End With
    End If
End Sub

Public Function GetPatientBaby(ByVal lngPatient As Long, lngPatientPage As Long) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 序号, 婴儿姓名 From 病人新生儿记录 Where 病人id = [1] And 主页ID = [2]"
    On Error GoTo errH
    Set GetPatientBaby = zlDatabase.OpenSQLRecord(strSQL, "读取新生儿记录", lngPatient, lngPatientPage)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim strValue As String
    
    On Error Resume Next
    '费用金额小数点位数
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '卡号显示方式
    gblnShowCard = zlDatabase.GetPara(12, glngSys) = "0"
    
   '就诊卡号长度
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbytCardNOLen = Val(Split(strValue, "|")(4))
    '问题:35242
    gbln简码切换 = IIf(Val(zlDatabase.GetPara("简码匹配方式切换", , , 1)) = 1, 1, 0) = 1
        
    '挂号有效天数
    '刘兴洪:34717
    '两位:前一位普能挂号;后一位急诊挂号
    strValue = zlDatabase.GetPara(21, glngSys, , "01") & "1"
    gSysPara.Sy_Reg.bytNODaysGeneral = Val(Left(strValue, 1))
    gSysPara.Sy_Reg.bytNoDayseMergency = Val(Right(strValue, 1))
    'If gSysPara.Sy_Reg.bytNODaysGeneral = 0 Then gSysPara.Sy_Reg.bytNODaysGeneral = 1
    ' If gSysPara.Sy_Reg.bytNoDayseMergency = 0 Then gSysPara.Sy_Reg.bytNoDayseMergency = 1
    '49501
    gSysPara.byt病人审核方式 = Val(zlDatabase.GetPara(185, glngSys, , "0"))
    gSysPara.bln未入科禁止记账 = Val(zlDatabase.GetPara(215, glngSys, , "0")) = 1 '51612
    '日报统计时间允许
    gblnDailyTime = zlDatabase.GetPara(22, glngSys) = "1"
    
    '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys))
    
    
    '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
    gstrMatchMode = zlDatabase.GetPara(44, glngSys)
                
    '刷卡要求输入密码
    gstrCardPass = zlDatabase.GetPara(46, glngSys, , "0000000000")
    gbln开单人 = zlDatabase.GetPara(52, glngSys) = "1"
    gbyt医保对码检查 = Val(zlDatabase.GetPara(59, glngSys))
    gcurMaxMoney = Val(zlDatabase.GetPara(60, glngSys))
    gbln报警包含划价费用 = zlDatabase.GetPara(98, glngSys) = "1"
    '刘兴洪 问题:????    日期:2010-12-06 23:38:53
    '费用单价保留位数
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    
    InitSysPar = True
End Function

Public Sub InitLocPar(bytUseType As Byte)
'功能：初始化费用本机参数
'参数：bytUseType=0-住院记帐,1-分散记帐,2-医技记帐,3-门诊,-1-设置
    Dim strValue As String
    
    If bytUseType = -1 Then Exit Sub
    
    '输入匹配类型
    gstrLike = IIf(zlDatabase.GetPara("输入匹配") = 0, "%", "")
    strValue = zlDatabase.GetPara("输入法")
    gstrIme = IIf(strValue = "", "不自动开启", strValue)
    gbytCode = Val(zlDatabase.GetPara("简码方式"))
    gblnMyStyle = zlDatabase.GetPara("使用个性化风格") = "1"

    
    '结帐可选择出院病人天数
    gintOutDay = Val(zlDatabase.GetPara("出院病人天数", glngSys, glngModul))
    
    gblnTime = zlDatabase.GetPara("变价数次", glngSys, glngModul) = "1"
    If bytUseType = 3 Then
        gstr收费类别 = zlDatabase.GetPara("收费类别", glngSys, 1121)
    Else
        gstr收费类别 = zlDatabase.GetPara("收费类别", glngSys, 1150)
        gbln护士 = zlDatabase.GetPara("显示护士", glngSys, glngModul) = "1"
    End If
    
    '留观病人记帐
    If bytUseType <> 3 Then
        gbln门诊留观 = (zlDatabase.GetPara("门诊留观病人记帐", glngSys, 1150, "0") = "1")
        gbln住院留观 = (zlDatabase.GetPara("住院留观病人记帐", glngSys, 1150, "0") = "1")
    End If
End Sub

Public Sub GetBillDeptID(strNO As String, lng开单ID As Long, lng执行ID As Long, ByVal int来源 As Integer)
'功能：获取一张记帐单据的开单科室和执行科室ID
'int来源:1-门诊;2-住院
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 开单部门ID,执行部门ID From " & IIf(int来源 = 1, "门诊费用记录", "住院费用记录") & " Where 记录性质=2 And 记录状态 IN(1,3) And NO=[1] And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlBill", strNO)
    If Not rsTmp.EOF Then
        lng开单ID = IIf(IsNull(rsTmp!开单部门ID), 0, rsTmp!开单部门ID)
        lng执行ID = IIf(IsNull(rsTmp!执行部门ID), 0, rsTmp!执行部门ID)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function GetBillTotal(objBill As ExpenseBill) As Currency
'功能：获取单据费目合计金额
    Dim objBillDetail As BillDetail
    Dim objBillIncome As BillInCome
    
    For Each objBillDetail In objBill.Details
        For Each objBillIncome In objBillDetail.InComes
            GetBillTotal = GetBillTotal + objBillIncome.实收金额
        Next
    Next
End Function
Public Function GetBillRowTotal(objBillInComes As BillInComes) As Currency
'功能：获取单据费目合计金额
    Dim objBillIncome As New BillInCome
    For Each objBillIncome In objBillInComes
        GetBillRowTotal = GetBillRowTotal + objBillIncome.实收金额
    Next
End Function

Public Function CheckScope(curL As Currency, curR As Currency, curI As Currency) As String
'功能：判断输入金额是否在原价和现从限定的范围内
'参数：curL=原价,curR=现价,curI=输入金额
'返回：如果不在范围内,则为提示信息,否则为空串
    If (curL >= 0 And curR >= 0) Or (curL <= 0 And curR <= 0) Then
        '如果数值符号相同,则用绝对值判断
        If Abs(curI) < Abs(curL) Or Abs(curI) > Abs(curR) Then
            CheckScope = "输入的金额绝对值不在范围(" & Format(Abs(curL), "0.00") & "-" & Format(Abs(curR), "0.00") & ")内."
        End If
    Else
        '如果符号不相同,则用原始范围判断
        If curI < curL Or curI > curR Then
            CheckScope = "输入的金额值不在范围(" & Format(curL, "0.00") & "-" & Format(curR, "0.00") & ")内."
        End If
    End If
End Function


Public Function Get收费执行科室ID(ByVal lng项目id As Long, ByVal int执行科室 As Integer, _
                ByVal lng病人科室ID As Long, ByVal lng开单科室ID As Long, Optional ByVal int范围 As Integer = 2, Optional ByVal lng病区ID As Long) As Long
'功能：获取非药收费项目的执行科室
'参数：int范围=1.门诊,2-住院

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim str药房 As String, lng药房 As Long
    Dim bytDay As Byte, strIDs As String
    
    On Error GoTo errH
    '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
    Select Case int执行科室
        Case 0 '0-无明确科室
            Get收费执行科室ID = UserInfo.部门ID
        Case 1 '1-病人所在科室
            Get收费执行科室ID = lng病人科室ID
        Case 2 '2-病人所在病区
            If int范围 = 1 Then
                Get收费执行科室ID = lng病人科室ID
            Else
                Get收费执行科室ID = lng病区ID
            End If
        Case 3 '3-操作员所在科室
            Get收费执行科室ID = UserInfo.部门ID
        Case 4 '4-指定科室
            strSQL = "" & _
            "   Select Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
            "   From 收费执行科室 A,部门表 C" & _
            "   Where A.收费细目ID=[1]　And A.执行科室ID+0=C.ID " & _
            "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
            "       And (A.病人来源 is NULL Or A.病人来源=[2])" & _
            "       And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
            " Order by Decode(A.病人来源,Null,2,1)" '默认科室优先
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng项目id, int范围, lng病人科室ID)
            If Not rsTmp.EOF Then
                '缺省取操作员所在科室
                rsTmp.Filter = "开单科室ID=" & UserInfo.部门ID
                If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
                If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0"
                If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
            End If
        Case 5 '院外执行(预留,程序暂未用)
        Case 6 '开单人科室
           Get收费执行科室ID = lng开单科室ID
    End Select
    If Get收费执行科室ID = 0 Then Get收费执行科室ID = UserInfo.部门ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckExecute(strNO As String, ByVal lng记帐单ID As Long, ByVal int来源 As Integer) As Boolean
'功能：判断费用对应的处方或摆药单是否已经审核
'参数：strNO   =费用单据号
'      记帐单ID=用于编辑的记帐单
'      int来源-1门诊;2住院
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Nvl(Count(ID),0) as 数目" & _
        " From " & IIf(int来源 = 1, "门诊费用记录", "住院费用记录") & _
        " Where NO=[1] And 记录性质=2 and 记帐单ID=[2]" & _
        " And 记录状态 IN(1,3) And 执行状态<>1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, lng记帐单ID)
    
    CheckExecute = (rsTmp!数目 = 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function PatiCanBilling(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
'功能：检查指定病人是否具有相关权限
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    PatiCanBilling = True
    Err = 0: On Error GoTo errH:
    If InStr(strPrivs, "出院未结强制记帐") > 0 _
        And InStr(strPrivs, "出院结清强制记帐") > 0 Then
        Exit Function
    End If
    
    strSQL = "Select A.姓名,B.出院日期,B.状态,X.费用余额" & _
        " From 病人信息 A,病案主页 B,病人余额 X" & _
        " Where A.病人ID=B.病人ID And A.病人ID=X.病人ID(+)" & _
        " And A.病人ID=[1] And B.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!出院日期) And Nvl(rsTmp!状态, 0) <> 3 Then Exit Function
        If InStr(strPrivs, "出院未结强制记帐") = 0 Then
            If Nvl(rsTmp!费用余额, 0) <> 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用未结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If InStr(strPrivs, "出院结清强制记帐") = 0 Then
            If Nvl(rsTmp!费用余额, 0) = 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用已结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetExamineItem(ByVal strItems As String, ByVal lngMediCareID As Long) As ADODB.Recordset
'功能:返回指定险类的收费项目要求审批的记录集
'参数:strItems-收费细目ID串,例如:"2369,2367,2368"
'     lngMediCareID-险类,例如:901
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select /*+ rule */ A.收费细目id" & vbNewLine & _
            "From 保险支付项目 A ,Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) B" & vbNewLine & _
            "Where A.险类 = [1] And A.要求审批 = 1 And A.收费细目id = B.Column_Value"
            
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngMediCareID, strItems)
    
    Set GetExamineItem = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRowByFeeItemID(ByRef ObjBillDetails As BillDetails, ByRef lngItemID As Long) As Long
'功能:根据收费项目ID返回其在单据中的行号,如果有重复的,只返回第一个
    Dim i As Long
    
    For i = 1 To ObjBillDetails.Count
        If lngItemID = ObjBillDetails(i).收费细目ID Then
            GetRowByFeeItemID = i: Exit Function
        End If
    Next
End Function

Public Function CheckExamine(ByRef ObjBillDetails As BillDetails, ByRef rsMedAudit As ADODB.Recordset, ByRef lngMediCareID As Long) As Boolean
'功能:根据给定的收费项目对象集和病人审批项目记录集检查相应的收费项目是否需要审批
    Dim i As Long, j As Long, strTmp As String
    Dim rsTmp As ADODB.Recordset
    
    For i = 1 To ObjBillDetails.Count
        strTmp = strTmp & "," & ObjBillDetails(i).收费细目ID
    Next
    Set rsTmp = GetExamineItem(Mid(strTmp, 2), lngMediCareID)
    
    strTmp = ""
    For i = 1 To rsTmp.RecordCount
        rsMedAudit.Filter = "项目ID=" & rsTmp!收费细目ID
        If rsMedAudit.RecordCount = 0 Then
            strTmp = strTmp & "," & GetRowByFeeItemID(ObjBillDetails, rsTmp!收费细目ID)
        ElseIf Not IsNull(rsMedAudit!可用数量) Then
            j = GetRowByFeeItemID(ObjBillDetails, rsTmp!收费细目ID)
            If ObjBillDetails(j).数次 > rsMedAudit!可用数量 Then
                MsgBox "第" & j & "行收费项目的数次超过了批准的使用限量" & rsMedAudit!可用数量 & ".", vbInformation, gstrSysName
                CheckExamine = False: Exit Function
            End If
        End If
        
        rsTmp.MoveNext
    Next
    
    If strTmp <> "" Then
        MsgBox "第" & Mid(strTmp, 2) & "行收费项目要求审批,当前病人未被批准使用!", vbInformation, gstrSysName
        CheckExamine = False: Exit Function
    End If
    CheckExamine = True
End Function


Public Function zlIsShowDeptCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查部门信息是否加载编码
    '返回:显示编码,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 13:11:01
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    If mlng部门编码平均长度 = 0 Then
        strSQL = "Select Avg(length(编码)) As 长度 From 部门表"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "取部门编码的平均长度")
        mlng部门编码平均长度 = Val(Nvl(rsTemp!长度))
    End If
    '由于编码长度可能过长,无法显示部门的名称,因此自动显示和不显示编码,当大于5时,不显示.小于5时,显示
   zlIsShowDeptCode = mlng部门编码平均长度 <= 5
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, ByVal strSearch As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:部门选择器
    '入参:cboDept-指定的部门部件
    '     rsDept-指定的部门
    '     strSearch-要搜索的串
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串,
    Dim strIDs As String
    
    
    
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编码) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编码)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编码) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!名称)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编码)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!名称))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID = 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
    
    '刘兴洪:直接定位
    If lngDeptID <> 0 Then GoTo GoOver:

    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = "简码"
    Case Else
        rsTemp.Sort = "编码"
    End Select
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "缺省", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboDept, lngDeptID, True
    zlCommFun.PressKey vbKeyTab
    zlSelectDept = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function




Public Function zlGetRegEventsCons(Optional strFieldName As String = "急诊", Optional strAliaName As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号项目的限制条件
    '入参:strFieldName-包含别外还字段(如急诊)
    '       strAliaName:别名
    '出参:
    '返回:返回条件
    '编制:刘兴洪
    '日期:2010-12-20 16:33:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strTimeName As String
    strFieldName = IIf(strAliaName <> "", strAliaName & ".", "") & strFieldName
    strTimeName = IIf(strAliaName <> "", strAliaName & ".", "") & " 登记时间"
    
    With gSysPara.Sy_Reg
        strWhere = ""
        If .bytNODaysGeneral <> 0 Or .bytNoDayseMergency <> 0 Then
            If .bytNODaysGeneral <> 0 Then
                strWhere = strWhere & " Or ( nvl(" & strFieldName & ",0)=0  And " & strTimeName & ">Trunc(Sysdate-" & .bytNODaysGeneral & "))"
            Else
                strWhere = strWhere & " Or  nvl(" & strFieldName & ",0)=0   "
            End If
            If .bytNoDayseMergency <> 0 Then
                strWhere = strWhere & " Or ( nvl(" & strFieldName & ",0)=1  And " & strTimeName & ">Trunc(Sysdate-" & .bytNoDayseMergency & "))"
            Else
                strWhere = strWhere & " Or nvl(" & strFieldName & ",0)=1  "
            End If
        End If
        If strWhere <> "" Then
            strWhere = " And  (" & Mid(strWhere, 4) & ")"
        End If
    End With
    zlGetRegEventsCons = strWhere
End Function
Public Function zlIsAllowFeeChange(lng病人ID As Long, lng主页ID As Long, _
   Optional int状态 As Integer = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许费用变动
    '入参:int状态-(-1表示从数据库中读取审核标志进行判断;>0表示,直接根据该状态进行判断)
    '返回:允许变动返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-05-21 15:44:47
    '问题:49501,51612
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    If gSysPara.byt病人审核方式 = 0 And gSysPara.bln未入科禁止记账 = False Then
        ''保持歉容
        zlIsAllowFeeChange = True: Exit Function
    End If
   
    strSQL = "" & _
    " Select Nvl(审核标志,0) as 审核标志,nvl(状态,0) as 状态" & _
    " From 病案主页 " & _
    " Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    
    If rsTemp.EOF Then
        MsgBox "未找到对应的病人信息,不允许进行记录操作!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '检查未入科病人不允许记账
    If gSysPara.bln未入科禁止记账 And Val(Nvl(rsTemp!状态)) = 1 Then
        '51612
        MsgBox "病人未入科(第" & lng主页ID & "次住院) ,不能对该病人进行记账或销账操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '审核相关检查
    If gSysPara.byt病人审核方式 = 0 Then zlIsAllowFeeChange = True: Exit Function
    If int状态 < 0 Then
        int状态 = Val(Nvl(rsTemp!审核标志))
    End If
    '检查相关状态
    If int状态 = 1 Then
        MsgBox "病人在第" & lng主页ID & "次住院中已经开始审核费用,不能对该病人进行费用变动。", vbInformation, gstrSysName
        Exit Function
    End If
    If int状态 = 2 Then
        MsgBox "已经完成了对病人第" & lng主页ID & "次住院费用的审核,不能对该病人进行费用变动。", vbInformation, gstrSysName
        Exit Function
    End If
    zlIsAllowFeeChange = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

