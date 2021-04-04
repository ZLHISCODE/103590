Attribute VB_Name = "MdlDrugStore"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gstrSQL As String
Public gobjBrower As Object

Public glngModul As Long
Public glngSys As Long                      '系统编号参数
Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrUnitName As String               '用户单位名称
Public gstrDbUser As String                 '当前用户姓名
Public gstrprivs As String                  '当前用户具有的当前模块的功能
Public gstrMatchMethod As String            '匹配方式:0表示双向匹配

Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称
Public gstrTryUse As String                 '试用否
Public gbytSimpleCodeTrans As Byte          '卡片界面是否允许简码切换控制

Public gobjCharge As Object                 '划价部件
Public gobjStuff As Object                  '卫材部件

Public Const gint门诊药房 As Integer = 2
Public Const gint住院药房 As Integer = 3

Public gobjESign As Object '电子签名接口
Public gblnESign处方发药 As Boolean         '处方发药场合是否启用
Public gblnESign部门发药 As Boolean         '部门发药场合是否启用
Public gblnESignUserStoped As Boolean       '用户电子签名证书是否停用

Public grsMaster As New ADODB.Recordset        '药品选择器：药品规格缓存数据集
Public grsMasterInput As New ADODB.Recordset   '药品选择器：药品规格录入简码时的缓存数据集
Public grsSlave As New ADODB.Recordset         '药品选择器：批次缓存数据集

Public gstrPriceClass As String         '价格等级

Public Enum EsignTache
    Dosage = 1  '配药
    send = 2    '发药
    returnStep = 3 '退药
End Enum

Public Const DblFrmHeight As Double = 3630

Public Const glngRowByFocus = &HFFE3C8
Public Const glngRowByNotFocus = &HF4F4EA
Public Const glngFixedForeColorByFocus = &HFF0000
Public Const glngFixedForeColorNotFocus = &H80000012

Public Const gstrUnit_DYEY = "大连医科大学附属第二医院"
Public Const gstrUnit_DLSY = "大连市第三人民医院"

'处方及部门发药字体颜色常数
Public Const glng退药 As Long = &HC0&
Public Const glng发药 As Long = &HC00000
Public Const glng正常 As Long = &H80000008
Public Const strAsc As String = "♂"                   '升序
Public Const strDesc As String = "♀"                  '降序

'LED显示相关变量
Public glngLEDModal As Long                'LED模块代码
Public grsLEDComponent As New ADODB.Recordset  'LED部件的数据库信息
Public gobjLEDShow As Object               'LED部件

'模块号
Public Enum 模块号
    外购入库 = 1300
    自制入库 = 1301
    其他入库 = 1302
    差价调整 = 1303
    药品移库 = 1304
    药品领用 = 1305
    其他出库 = 1306
    药品盘点 = 1307
    药品计划 = 1330
    质量管理 = 1331
End Enum


'用户信息------------------------
Public Type TYPE_USER_INFO
    用户ID As Long
    用户编码 As String
    用户姓名 As String
    用户简码 As String
    部门ID As Long
    部门编码 As String
    部门名称 As String
    strMaterial As String
End Type
Public UserInfo As TYPE_USER_INFO

'部门发药中各种颜色设置
Public Enum mListColor
    LowerLimit = &HC000C0                       '低于库存下限：紫色
    SumTotal = vbBlue                           '小计、合计：蓝色
    State_Send = &HFFDDDD                       '发药状态：浅蓝色
    State_UnProcess = &H80000005                '不处理状态：白色
    State_Reject = &HDBDBDB                     '拒发状态：浅灰色
    State_Shortage = &HD7D7FF                   '缺药状态：浅红色
    State_RejectRestore = &HD7D7FF              '拒发恢复状态：浅红色
    State_RejectUnProcess = &H80000005          '拒发不处理状态：白色
    Return_Original = &H80000008                '退药状态（原始单据）：黑色
    Return_Sended = &HC00000                    '退药状态（已发药单据）：蓝色
    Return_Returned = &HC0&                     '退药状态（退药单据）：红色
End Enum

'药房模块要使用到的系统参数
Public Type Type_SysParms
    P6_未审核记帐处方发药 As Integer
    P9_费用金额保留位数 As Integer
    P15_门诊收费与发药分离 As Integer
    P16_住院记帐与发药分离 As Integer
    P23_已结帐单据操作 As Integer
    P25_使用电子签名 As Integer
    P26_电子签名场合 As String
    P28_门诊病人消费时需要刷卡验证 As String
    P29_指导批发价定价单位 As Integer
    P44_输入匹配 As String
    P54_时价药品以加价率入库 As Integer
    P64_审核限制 As Integer
    P68_门诊药嘱先作废后退药 As Integer
    P70_过敏登记有效天数 As Integer
    P75_外购入库需要核查 As Integer
    P76_时价药品直接确定售价 As Integer
    P85_药房查看单据成本价 As Integer
    P96_药品填单下可用库存 As Integer
    P98_记帐报警包含划价费用 As Integer
    P126_时价药品售价加成方式 As Integer
    P148_未收费处方发药 As Integer
    P149_效期显示方式 As Integer
    P150_药品出库优先算法 As Integer
    P153_配置中心 As Long
    P163_项目执行前必须先收费或先记帐审核 As Integer
    Para_输入方式 As String
    P214_首次医嘱执行需要审核  As Integer
    P221_药品结存时点 As Integer
    P174_药品移库明确批次 As Integer
    P175_药品领用明确批次 As Integer
    P222_药房自动化发药接口 As Integer
    P240_药房处方审查 As Integer
    P241_处方审查时机 As Integer
    P275_零差价管理模式 As Integer
    P213_中药配方每行中药味数 As Integer
    
End Type
Public gtype_UserSysParms As Type_SysParms     '系统参数

'公共模块参数
Public gstrLike As String                       '输入匹配
Public gblnMyStyle As Boolean                   '个性化风格

Public gint简码方式 As Integer              '0-拼音，1-五笔
Public gint药品名称显示 As Integer          '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
Public gint输入药品显示 As Integer          '0-按输入匹配显示，1-固定显示通用名和商品名

'业务单据号
Public Enum 单据号
    外购入库 = 1
    自制入库 = 2
    协药入库 = 3
    其他入库 = 4
    差价调整 = 5
    药品移库 = 6
    药品领用 = 7
    收费处方发药 = 8
    记帐单处方发药 = 9
    记帐表处方发药 = 10
    其他出库 = 11
    盘点表 = 12
    调价变动 = 13
    盘点单 = 14
    留存记录 = 27
End Enum

'私有、公共模块参数
Public Enum 参数_协定入库_私有
    P1_是否选择库房 = 1
    P2_存盘打印 = 2
    P3_审核打印 = 3
End Enum

Public Enum 参数_药品申领_私有
    P1_是否选择库房 = 1
    P2_药品单位 = 2
    P3_排序 = 3
    P4_存盘打印 = 4
    P5_审核打印 = 5
    P6_查询天数 = 6
End Enum

Public Enum 参数_处方发药_私有
    P1_列设置 = 1
    P2_字体 = 2
End Enum

Public Enum 参数_处方发药_公共
    P1_收费处方显示方式 = 1
    P2_记帐处方显示方式 = 2
    P3_查询天数 = 3
    P4_处方颜色 = 4
    P5_打印包含记帐单 = 5
    P6_打印退费单据间隔 = 6
    P7_打印延迟 = 7
    P8_显示病区处方 = 8
    P9_刷新间隔 = 9
    P10_校验发药人 = 10
    P11_校验方式 = 11
    P12_校验配药人 = 12
    P13_自动销帐 = 13
End Enum

Public Enum 参数_部门发药_私有
    P1_列设置 = 1
    P2_字体 = 2
End Enum

Public Enum 参数_部门发药_公共
    P1_查询天数 = 1
    P2_发药规则 = 2
    P3_简要条件 = 3
    P4_领药人签名 = 4
    P5_缺药检查 = 5
    P6_退药人签名 = 6
    P7_病区发药方式 = 7
    P8_自动刷新未发药清单 = 8
End Enum

Public Enum 参数_处方审查_公共
    P1_审查标准 = 1
End Enum

'药品金额、价格、数量最大精度
Public Type Type_Digits
    Digit_金额 As Integer
    Digit_成本价 As Integer
    Digit_零售价 As Integer
    Digit_数量 As Integer
End Type
Public gtype_UserDrugDigits As Type_Digits

Public Type Type_SaleDigits
    Digit_成本价 As Integer
    Digit_零售价 As Integer
    Digit_数量 As Integer
End Type
Public gtype_UserSaleDigits As Type_SaleDigits

'单据操作控制
Private Type Type_BillControl
    bln是否控制 As Boolean
    int时间限制 As Integer
    bln他人单据 As Boolean
    dbl金额上限 As Double
End Type
Private gtype_myBillControl As Type_BillControl


'处方类型名称，按顺序，用;分隔
Public Const gconstrRecipeType = "普通;儿科;急诊;精二;精一;麻醉"

'默认处方颜色：普通－白色；急诊－淡黄色；儿科－淡绿色；麻醉、精一－淡红色；精二－白色
Private Const gconlng普通 = &HFFFFFF
Private Const gconlng儿科 = &HC0FFC0
Private Const gconlng急诊 = &HC0FFFF
Private Const gconlng精二 = &HFFFFFF
Private Const gconlng精一 = &HC0C0FF
Private Const gconlng麻醉 = &HC0C0FF

Public Type InOutType
    bln外购入库 As Boolean
    bln自制入库 As Boolean
    bln协药入库 As Boolean
    bln其它入库 As Boolean
    bln差价调整 As Boolean
    bln药品移库 As Boolean
    bln药品领用 As Boolean
    bln收费处方发药 As Boolean
    bln记帐单处方发药 As Boolean
    bln记帐表处方发药 As Boolean
    bln其他出库 As Boolean
    bln盘点表 As Boolean
    bln调价变动 As Boolean
    bln盘点单 As Boolean
    bln药品留存 As Boolean
End Type
Public gInOutType As InOutType

Public Enum 医院业务
    support门诊预算 = 0
    support门诊退费 = 1
    support预交退个人帐户 = 2
    
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
    support出院无实际交易 = 29       '出院接口中是否要与接口商进行交易
    support允许部分冲销明细 = 32    '允许针对住院记帐处方的每笔明细进行部分冲销
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
    support住院结算作废 = 34        'HIS始终认为住院支持结算作废，如果不支持需医保接口内部处理，返回假即可；增加该参数是为了配合GetCapability交易来检查各种结算方式是否支持全退
    support负数记帐 = 35            '是否允许负数记帐，操作员首先要拥有负数记帐的权限。此参数缺省为真，不支持的接口需单独处理
    support结帐_指定住院次数 = 36   '是否支持指定住院次数进行医保结算
    support结帐_指定日期范围 = 37   '是否支持指定结帐日期范围进行医保结算
    support结帐_设置婴儿费条件 = 38 '是否允许设置婴儿费条件
    
    support门诊结帐 = 41            '是否支持门诊医保病人的记帐费用使用门诊结帐来完成
End Enum

Public Sub setNOtExcetePrice()
    '将到时间还未执行调价药品执行调价
    Dim rstemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct i.Id As 药品id " & _
               " From 收费项目目录 I, 收费价目 N, 药品规格 P" & _
               " Where i.Id = n.收费细目id And i.Id = p.药品id And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & _
                   " n.变动原因 = 0 And Sysdate>n.执行日期" & GetPriceClassString("N") & _
               " Union " & _
               " Select Distinct a.药品id From 药品价格记录 A Where a.记录状态 = 0 And a.执行日期 <= Sysdate " & _
               " Order By 药品id "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "执行调价")
    
    If rstemp.RecordCount = 0 Then Exit Sub
    
    For i = 0 To rstemp.RecordCount - 1
        gstrSQL = "Zl_药品收发记录_Adjust(" & rstemp!药品ID & ")"
        zldatabase.ExecuteProcedure gstrSQL, "执行调价"
        rstemp.MoveNext
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function CheckPriceAdjustByNO(ByVal Int单据 As Integer, ByVal lng库房id As Long, ByVal strNo As String, Optional ByVal lng检查库房id As Long = 0) As Boolean
    '按单据号检查零差价
    Dim rsData As ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '如果没开启全局的零差价管理，则不进行后续检查，返回true
    If Val(zldatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjustByNO = True: Exit Function
    
    gstrSQL = "Select 药品id, Nvl(批次, 0) As 批次 From 药品收发记录 " & _
        " Where 单据 = [1] And 库房id = [2] And NO = [3] And 审核日期 Is Null"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjustByNO", Int单据, lng库房id, strNo)
    
    If rsData.EOF Then CheckPriceAdjustByNO = True: Exit Function
    
    Do While Not rsData.EOF
        If IsPriceAdjustMod(Val(rsData!药品ID)) = True Then
            If CheckPriceAdjust(rsData!药品ID, IIf(lng检查库房id = 0, lng库房id, lng检查库房id), rsData!批次) = False Then
                gstrSQL = "Select '[' || 编码 || ']' || 名称 || '(' || 规格 || ')' as 药品 From 收费项目目录 Where ID = [1]"
                Set rsItem = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjustByNO", Val(rsData!药品ID))
                If Not rsItem.EOF Then
                    MsgBox "单据[" & strNo & "]中的药品 " & rsItem!药品 & "已启用了零差价管理，但在药房中售价和成本价不一致，不能发药 ！" & _
                     vbCrLf & "请先进行调价再发药。", vbInformation, gstrSysName
                Else
                    MsgBox "单据[" & strNo & "]中的药品已启用了零差价管理，但在药房中售价和成本价不一致，不能发药 ！" & _
                     vbCrLf & "请先进行调价再发药。", vbInformation, gstrSysName
                End If
                
                Exit Function
            End If
        End If
        
        rsData.MoveNext
    Loop
    
    CheckPriceAdjustByNO = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPriceAdjust(ByVal lng药品id As Long, ByVal lng库房id As Long, ByVal lng批次 As Long) As Boolean
    '零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
    '定价药品：售价是固定的，比较所有药房的成本价，如果存在不一致的就不能销售出库
    '时价药品：比较药房库存记录的零售价和成本价，如果存在不一致的就不能销售出库
    '无库存时：成本价取药品规格的成本价
    '参数：lng药品id-药品规格ID，为0则检查所有药品；lng库房id-对应的库房ID，为0则检查所有库房；lng批次-对应的批次，如果传入-1则不关联批次
    '返回：True-正常；false-有不满足零差价管理要求的药品
    '
    Dim rsData As ADODB.Recordset
    Dim str条件 As String
    
    On Error GoTo errHandle
    
    '如果没开启全局的零差价管理，则不进行后续检查，返回true
    If Val(zldatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjust = True: Exit Function
    
    '检查有无库存
    If lng药品id > 0 Then
        If lng库房id > 0 Then
            gstrSQL = "Select 1 from 药品库存 Where 性质=1 and 药品id=[1] and 库房id=[2] " & _
                " And Not (nvl(批次,0) = 0 And 可用数量 < 0 And 实际数量 = 0 And 实际金额 = 0 And 实际差价 = 0)"
            
            If lng批次 > 0 Then
                gstrSQL = gstrSQL & " and Nvl(批次,0)=[3] "
            End If
        Else
            gstrSQL = "Select 1 from 药品库存 Where 性质=1 and 药品id=[1] " & _
                " And Not (nvl(批次,0) = 0 And 可用数量 < 0 And 实际数量 = 0 And 实际金额 = 0 And 实际差价 = 0)"
        End If
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjust", lng药品id, lng库房id, lng批次)
        
        If rsData.EOF Then
            '无库存时，从收费价目取售价，从药品规格取成本价，并比较价格
            gstrSQL = "Select a.成本价, b.现价 As 售价 " & _
                " From 药品规格 A, 收费价目 B " & _
                " Where a.药品id = b.收费细目id And (Sysdate Between b.执行日期 And b.终止日期) And Nvl(a.是否零差价管理, 0) = 1 " & _
                " And b.现价 <> a.成本价 And a.药品id = [1] " & GetPriceClassString("B")
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjust", lng药品id)
            
            If rsData.EOF Then
                '没找到表示价格一致
                CheckPriceAdjust = True
            Else
                '找到表示价格不一致
                CheckPriceAdjust = False
            End If
            
            Exit Function
        End If
    End If
    
    If lng药品id > 0 Then
        str条件 = IIf(str条件 = "", "", str条件) & " and a.药品id=[1] "
    End If
    
    If lng库房id > 0 Then
        str条件 = IIf(str条件 = "", "", str条件) & " and d.库房id=[2] "
    End If
    
    If lng批次 >= 0 Then
        str条件 = IIf(str条件 = "", "", str条件) & " and nvl(d.批次,0)=[3] "
    End If
    
    gstrSQL = "Select a.药品id, '['|| c.编码 || ']'|| c.名称||decode(c.产地,null,null,'('||c.产地||')') ||c.规格 As 通用名 " & vbNewLine & _
        "       From 药品规格 A, 收费价目 B, 收费项目目录 C, 药品库存 D" & vbNewLine & _
        "       Where a.药品id = b.收费细目id And a.药品id = c.Id And a.药品id = d.药品id And d.性质 = 1 And (Sysdate Between b.执行日期 And b.终止日期) And" & vbNewLine & _
        "             (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.是否变价 = 0 And Nvl(a.是否零差价管理, 0) = 1 And" & vbNewLine & _
        "             b.现价 <> nvl(d.平均成本价,a.成本价) " & str条件 & GetPriceClassString("B") & vbNewLine & _
        "  And Not (nvl(D.批次,0) = 0 And D.可用数量 < 0 And D.实际数量 = 0 And D.实际金额 = 0 And D.实际差价 = 0) " & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select a.药品id, '['|| c.编码 || ']'|| c.名称||decode(c.产地,null,null,'('||c.产地||')') ||c.规格 As 通用名 " & vbNewLine & _
        " From 药品规格 A, 收费项目目录 C, 药品库存 D, 部门表 E" & vbNewLine & _
        " Where a.药品id = c.Id And a.药品id = d.药品id And d.库房id = e.Id And d.性质 = 1 And c.是否变价 = 1 And" & vbNewLine & _
        "      (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.是否零差价管理, 0) = 1 And nvl(d.零售价,0) <> nvl(d.平均成本价,a.成本价) " & str条件 & _
        "  And Not (nvl(D.批次,0) = 0 And D.可用数量 < 0 And D.实际数量 = 0 And D.实际金额 = 0 And D.实际差价 = 0) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjust", lng药品id, lng库房id, lng批次)
    
    '没找到不满足零差价管理要求的记录，返回true
    If rsData.EOF Then CheckPriceAdjust = True: Exit Function
    
    CheckPriceAdjust = False
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPriceAdjustBatch(ByVal lng库房id As Long, ByVal str药品批次 As String) As String
    '零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致），批量查询模式
    '定价药品：售价是固定的，比较所有药房的成本价，如果存在不一致的就不能销售出库
    '时价药品：比较药房库存记录的零售价和成本价，如果存在不一致的就不能销售出库
    '无库存时：成本价取药品规格的成本价
    '参数：lng库房id-对应的库房ID；str药品批次-对应的药品ID及批次，格式为“药品ID,批次|药品ID,批次...”
    '返回：空串-正常；非空串-有不满足零差价管理要求的药品名称
    '
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    Dim strInfo As String
    
    On Error GoTo errHandle
    
    '如果没开启全局的零差价管理，则不进行后续检查，返回true
    If Val(zldatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjustBatch = "": Exit Function

    gstrSQL = "Select Distinct a.药品id, '[' || c.编码 || ']' || c.名称 || Decode(c.产地, Null, Null, '(' || c.产地 || ')') || c.规格 As 通用名" & vbNewLine & _
        "From 药品规格 A, 收费价目 B, 收费项目目录 C, 药品库存 D, Table(f_Num2list2([2], '|', ',')) T" & vbNewLine & _
        "Where a.药品id = b.收费细目id And a.药品id = c.Id And a.药品id = d.药品id And d.性质 = 1 And (Sysdate Between b.执行日期 And b.终止日期) And" & vbNewLine & _
        "      (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.是否变价 = 0 And Nvl(a.是否零差价管理, 0) = 1 And" & vbNewLine & _
        "      b.现价 <> Nvl(d.平均成本价, a.成本价) And a.药品id = t.C1 And d.库房id = [1] And Nvl(d.批次, 0) = t.C2 And b.价格等级 Is Null" & vbNewLine & _
        " And Not (nvl(D.批次,0) = 0 And D.可用数量 < 0 And D.实际数量 = 0 And D.实际金额 = 0 And D.实际差价 = 0) " & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select Distinct a.药品id, '[' || c.编码 || ']' || c.名称 || Decode(c.产地, Null, Null, '(' || c.产地 || ')') || c.规格 As 通用名" & vbNewLine & _
        "From 药品规格 A, 收费项目目录 C, 药品库存 D, 部门表 E, Table(f_Num2list2([2], '|', ',')) T" & vbNewLine & _
        "Where a.药品id = c.Id And a.药品id = d.药品id And d.库房id = e.Id And d.性质 = 1 And c.是否变价 = 1 And" & vbNewLine & _
        "      (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.是否零差价管理, 0) = 1 And" & vbNewLine & _
        "      Nvl(d.零售价, 0) <> Nvl(d.平均成本价, a.成本价) And a.药品id = t.C1 And d.库房id = [1] And Nvl(d.批次, 0) = t.C2 " & vbNewLine & _
        " And Not (nvl(D.批次,0) = 0 And D.可用数量 < 0 And D.实际数量 = 0 And D.实际金额 = 0 And D.实际差价 = 0) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjustBatch", lng库房id, str药品批次)

    '没找到不满足零差价管理要求的记录，返回true
    If rsData.EOF Then CheckPriceAdjustBatch = "": Exit Function
    
    For i = 1 To rsData.RecordCount
        If i > 3 Then
            Exit For
        End If
         
        strInfo = IIf(strInfo = "", "", strInfo & vbCrLf) & rsData!通用名
        
        rsData.MoveNext
    Next

    CheckPriceAdjustBatch = strInfo
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsPriceAdjustMod(ByVal lng药品id As Long) As Boolean
    '判断药品是否启用零差价管理
    Dim rsData As ADODB.Recordset
    
    gstrSQL = "Select Nvl(是否零差价管理, 0) As 是否零差价管理 From 药品规格 Where 药品id = [1] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "IsPriceAdjustMod", lng药品id)
    
    If rsData.EOF Then IsPriceAdjustMod = False: Exit Function
    
    IsPriceAdjustMod = (rsData!是否零差价管理 = 1)
End Function
Public Function GetMediPackerDetail(ByVal lng收发ID As Long, ByVal str剂型 As String, ByVal str类型 As String) As String
    '用于药品分包机接口
    '传入收发ID；返回要传入分包机接口的明细字符串
    '返回的字符串按一定顺序和格式
    
    Dim rsData As ADODB.Recordset
    Dim rsGetNext As ADODB.Recordset
    Dim n As Integer
    Dim strReturn As String
    Dim strLastTime As String
    Dim intCount As Integer
    Dim blnErr As Boolean
    
     gstrSQL = "Select /*+ Rule */ A.收发id, A.住院号, A.病人id, A.姓名, A.病区编码, A.病区名称, A.开单人, A.床号, A.用法, A.药品编码, A.药品名称, A.规格, A.剂量系数, A.剂量单位, A.服用数量,A.审核人," & _
        " A.首次时间, A.末次时间,A.开始执行时间, A.频率间隔, A.间隔单位, A.执行时间方案, Nvl(B.发送数次, 0) As 次数, A.发药数量,整包装 " & _
        " From (Select Distinct A.ID As 收发id, B.标识号 As 住院号, B.病人id, B.姓名, C.编码 As 病区编码, C.名称 As 病区名称, B.开单人, B.床号, A.用法,A.审核人," & _
        " D.编码 As 药品编码, D.名称 As 药品名称, D.规格, E.剂量系数, F.计算单位 As 剂量单位, H.单次用量 / E.剂量系数 As 服用数量, G.首次时间, G.末次时间," & _
        " H.开始执行时间 , H.频率间隔, H.间隔单位, H.执行时间方案, H.相关id, G.发送号, A.实际数量 * Nvl(A.付数, 1) / E.住院包装 As 发药数量,decode(mod(A.实际数量 * Nvl(A.付数, 1) , E.药库包装),0,1,0) 整包装 " & _
        " From 药品收发记录 A, 住院费用记录 B, 部门表 C, 收费项目目录 D, 药品规格 E, 诊疗项目目录 F, 病人医嘱发送 G, 病人医嘱记录 H "
    If str剂型 <> "所有" Then
        gstrSQL = gstrSQL & " ,药品特性 I, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) J "
    End If
    gstrSQL = gstrSQL & " Where A.费用id = B.ID And B.病人病区id = C.ID And A.药品id = D.ID And A.药品id = E.药品id And E.药名id = F.ID And " & _
        " B.医嘱序号 = G.医嘱id And B.NO = G.NO And B.医嘱序号 = H.ID And A.ID = [1] "
    If str剂型 <> "所有" Then
        gstrSQL = gstrSQL & "And E.药名id = I.药名id And I.药品剂型 = J.Column_Value "
    End If
    gstrSQL = gstrSQL & ") A, 病人医嘱发送 B " & _
        " Where A.相关id = B.医嘱id(+) And A.发送号 = B.发送号(+) "

    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "医嘱明细", lng收发ID, str剂型)
    
    If rsData.RecordCount = 0 Then Exit Function
    
    With rsData
        If Not .EOF Then
            '如果是临嘱并且是整包装数量，则不发送到包药机
            If !整包装 = 0 Or str类型 = "长嘱" Then
                If Val(Nvl(!频率间隔, 0)) = 0 Or Nvl(!间隔单位, "") = "" Or Nvl(!执行时间方案, "") = "" Or str类型 = "临嘱" Then
                    intCount = 1
                Else
                    intCount = Val(!次数)
                    If intCount = 0 Then
                        gstrSQL = "Select Zl_Gettransexenumber([1],[2],[3],[4],[5],[6]) From Dual "
                        Set rsGetNext = zldatabase.OpenSQLRecord(gstrSQL, "取下次执行时间", CDate(!开始执行时间), CDate(!首次时间), CDate(!末次时间), Val(!频率间隔), !间隔单位, !执行时间方案)
                        If Not rsGetNext.EOF Then
                            intCount = Val(rsGetNext.Fields(0).Value)
                        End If
                    End If
                    If intCount = 0 Then
                        intCount = 1
                        blnErr = True
                    End If
                End If
                
                For n = 1 To intCount
                    strReturn = IIf(strReturn = "", "", strReturn & "|")
                    strReturn = strReturn & !收发ID
                    strReturn = strReturn & ";" & !住院号
                    strReturn = strReturn & ";" & !病人ID
                    strReturn = strReturn & ";" & Replace(Replace(!姓名, ";", ""), "|", "")
                    strReturn = strReturn & ";" & !病区编码
                    strReturn = strReturn & ";" & Replace(Replace(!病区名称, ";", ""), "|", "")
                    strReturn = strReturn & ";" & Replace(Replace(!开单人, ";", ""), "|", "")
                    strReturn = strReturn & ";" & Replace(Replace(Nvl(!床号, ""), ";", ""), "|", "")
                    strReturn = strReturn & ";" & Replace(Replace(Nvl(!用法, ""), ";", ""), "|", "")
                    strReturn = strReturn & ";" & ""    '服用时间说明
                    strReturn = strReturn & ";" & !药品编码
                    strReturn = strReturn & ";" & Replace(Replace(!药品名称, ";", ""), "|", "")
                    strReturn = strReturn & ";" & Replace(Replace(!规格, ";", ""), "|", "")
                    strReturn = strReturn & ";" & !剂量系数
                    strReturn = strReturn & ";" & !剂量单位
                    
                    If str类型 = "临嘱" Then
                        strReturn = strReturn & ";" & !发药数量
                    Else
                        strReturn = strReturn & ";" & IIf(blnErr = False, !服用数量, !发药数量)
                    End If
                    
                    If n = 1 Then
                        strLastTime = Format(!首次时间, "YYYY-MM-DD HH:MM:SS")
                    Else
                        gstrSQL = "Select Zl_Gettransexetime([1],[2],[3],[4],[5]) From Dual "
                        Set rsGetNext = zldatabase.OpenSQLRecord(gstrSQL, "取下次执行时间", CDate(!开始执行时间), CDate(strLastTime), Val(!频率间隔), !间隔单位, !执行时间方案)
                        If Not rsGetNext.EOF Then
                            strLastTime = Format(rsGetNext.Fields(0).Value, "YYYY-MM-DD HH:MM:SS")
                        End If
                    End If
                    
                    strReturn = strReturn & ";" & strLastTime
                    strReturn = strReturn & ";" & "1"           '分包设备编号
                    strReturn = strReturn & ";" & "0"           '优先标记
                    
                    If str类型 = "临嘱" Then
                        strReturn = strReturn & ";" & "1"
                    Else
                        strReturn = strReturn & ";" & "0"
                    End If
                    
                    strReturn = strReturn & ";" & !审核人
                Next
            End If
        End If
    End With
    
    GetMediPackerDetail = strReturn
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub OutPutData(ByVal strOutput As String)
    '用于编译后的用户环境调试，在其他调试或不方便查找问题时使用
    '将程序执行的关键流程，数据输出到外部日志文件，以此方便查找问题
    '注意：在需要调试时手工创建指定的日志文件，编译后环境时放到导航台程序所在目录，源代码环境时放到工程文件所在目录
    '注意：如果不需要调试了要及时删除日志文件，否则日志文件可能会逐步增大，特别是用户环境可能数据增长较快
    '各系统可以指定不同的日志文件名
    '日志内容自定义，参考格式：时间+程序内部过程/函数+业务流程/步骤+关键数据
    '默认的处理都加上时间，如果不需要可以去掉
    Dim objFile As New FileSystemObject
    Dim objTarget As TextStream
    Const STR_CONS_FILENAME As String = "zlDrugStore.log"
    
    err = 0
    
    On Error Resume Next
    
    '检查文件是否存在
    Set objTarget = objFile.OpenTextFile(App.Path & "\" & STR_CONS_FILENAME)
    
    '如果不存在则不输出内容
    If objTarget Is Nothing Then Exit Sub
    
'    If err <> 0 Then
'        '创建目标文件
'        Set objFile = CreateObject("Scripting.FileSystemObject")
'        Set objTarget = objFile.CreateTextFile(App.Path & "\" & STR_CONS_FILENAME, True)
'        objTarget.Close
'    End If
    
    err.Clear
    On Error GoTo ErrHand
    
    Open App.Path & "\" & STR_CONS_FILENAME For Append Shared As #1
    
    Print #1, Now & vbCrLf & strOutput
    Close #1
    
    Exit Sub
ErrHand:
    Close #1
'    MsgBox err.Description, vbExclamation + vbOKOnly
End Sub

Public Sub AutoAdjustPrice_ByID(ByVal lngDrugID As Long)
    '检查所有已到执行日期而价格未执行的药品，执行调价过程
    '按指定药品ID检查
    '在药品选择器中调用
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    
    On Error GoTo errHandle

    gstrSQL = "zl_药品收发记录_Adjust(" & lngDrugID & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice_ByID")

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckNotVerifyClosingAccount() As ADODB.Recordset
    '查询当前操作员所属的部门是否存在未审核的结存记录
    Dim rsData As ADODB.Recordset
    Dim strDept As String
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct b.Id, b.名称, '未审核误差' As 类型" & vbNewLine & _
            "From 部门人员 A, 部门表 B, 部门性质说明 C, 药品结存记录 D, 药品结存误差 E" & vbNewLine & _
            "Where a.部门id = b.Id And b.Id = c.部门id And b.Id = d.库房id And d.Id = e.结存id And a.人员id = [1] And" & vbNewLine & _
            "      c.工作性质 In ('西药库', '成药库', '中药库', '西药房', '成药房', '中药房', '制剂室') And d.审核日期 Is Null" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select Distinct b.Id, b.名称, '未审核结存' As 类型" & vbNewLine & _
            "From 部门人员 A, 部门表 B, 部门性质说明 C" & vbNewLine & _
            "Where a.部门id = b.Id And b.Id = c.部门id And a.人员id = [1] And c.工作性质 In ('西药库', '成药库', '中药库', '西药房', '成药房', '中药房', '制剂室') And" & vbNewLine & _
            "      Exists (Select 1 From 药品结存记录 D Where b.Id = d.库房id And d.审核日期 Is Null)"

    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "结存查询", UserInfo.用户ID)
    
    Set CheckNotVerifyClosingAccount = rsData
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AutoAdjustPrice_ByNO(ByVal Int单据 As Integer, ByVal strNo As String)
    '检查所有已到执行日期而价格未执行的药品，执行调价过程
    '按指定单据,NO中的药品号进行检查
    '个流通业务模块的审核时调用
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    Dim blnMore As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct b.药品id " & _
        " From 收费价目 A, 药品收发记录 B, 收费项目目录 C " & _
        " Where a.收费细目id = b.药品id And a.收费细目id = c.Id And Nvl(c.是否变价, 0) = 0 And a.变动原因 = 0 And a.执行日期 <= Sysdate And b.审核日期 Is Null " & _
        " And b.单据 = [1] And b.No = [2]" & GetPriceClassString("A") & _
        " Union " & _
        " Select Distinct a.药品id " & _
        " From 药品价格记录 A, 药品收发记录 B " & _
        " Where a.药品id = b.药品id And a.记录状态 = 0 And a.执行日期 <= Sysdate And b.审核日期 Is Null And " & _
        " b.单据 = [1] And b.No = [2] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "AutoAdjustPrice", Int单据, strNo)

    With rsData
        If .RecordCount > 20 Then
            blnMore = True
            Call FS.ShowFlash("正在批量执行调价，请稍后......")
        End If
        
        Do While Not .EOF
            lngAdjustID = !药品ID
            gstrSQL = "zl_药品收发记录_Adjust(" & lngAdjustID & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice")
            
            .MoveNext
        Loop
        
        If blnMore = True Then
            Call FS.StopFlash
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub AutoAdjustPrice_Batch()
    '检查所有已到执行日期而价格未执行的药品，执行调价过程
    '检查所有药品
    '在药品选择器数据集初始化时调用
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    Dim blnMore As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct a.收费细目id As 药品id" & vbNewLine & _
        "From 收费价目 A, 收费项目目录 B" & vbNewLine & _
        "Where a.收费细目id = b.Id And b.类别 In ('5', '6', '7') And Nvl(b.是否变价, 0) = 0 And a.变动原因 = 0 " & _
        "And a.执行日期 <= Sysdate" & GetPriceClassString("A") & vbNewLine & _
        "Union" & vbNewLine & _
        "Select Distinct a.药品id From 药品价格记录 A Where a.记录状态 = 0 And a.执行日期 <= Sysdate"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "AutoAdjustPrice")

    With rsData
        If .RecordCount > 20 Then
            blnMore = True
            Call FS.ShowFlash("正在批量执行调价，请稍后......")
        End If
        
        Do While Not .EOF
            lngAdjustID = !药品ID
            gstrSQL = "zl_药品收发记录_Adjust(" & lngAdjustID & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice")
            
            .MoveNext
        Loop
        
        If blnMore = True Then
            Call FS.StopFlash
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function Get成本价(ByVal lng药品id As Long, ByVal lng库房id As Long, ByVal lng批次 As Long) As Double
'功能：获取当前药品的成本价格
'参数：药品id,库房id,批次
'返回值： 成本价格
    Dim rsData As ADODB.Recordset
    Dim blnNullPrice As Boolean
    
    On Error GoTo errHandle
    
    gstrSQL = "select 平均成本价 from 药品库存 where 性质=1 and 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "成本价", lng药品id, lng库房id, lng批次)
    
    If rsData.EOF Then
        blnNullPrice = True
    ElseIf IsNull(rsData!平均成本价) = True Then
        blnNullPrice = True
    ElseIf Val(rsData!平均成本价) < 0 Then
        blnNullPrice = True
    End If
    
    If Not blnNullPrice Then
        Get成本价 = rsData!平均成本价
    Else
        '如果无法从库存中取成本价，则从药品规格中取
        gstrSQL = "select 成本价 from 药品规格 where 药品id=[1]"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "成本价", lng药品id)
        If Not rsData.EOF Then
            If Val(Nvl(rsData!成本价, 0)) > 0 Then
                Get成本价 = rsData!成本价
            End If
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get售价(ByVal bln是否时价 As Boolean, lng药品id As Long, ByVal lng库房id As Long, ByVal lng批次 As Long) As Double
    '功能：获取原始的售价单位售价，主要用于出库
    '参数: bln是否时价:false-定价,true-时价
    '返回值：最小单位的价格
    Dim rsData As ADODB.Recordset
    Dim dbl零售价 As Double, dbl指导零售价 As Double, dbl差价让利比 As Double, dbl加成率 As Double
    Dim dbl成本价 As Double
    
    On Error GoTo errHandle

    '取定价药品售价
    If bln是否时价 = False Then
        gstrSQL = "Select 现价 " & _
            " From 收费价目 A, 药品规格 B " & _
            " Where A.收费细目id = B.药品id And A.收费细目ID=[1] And Sysdate Between A.执行日期 And Nvl(A.终止日期,Sysdate) " & GetPriceClassString("A")
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "Get售价-取定价药品售价", lng药品id)
        
        If Not rsData.EOF Then
            Get售价 = rsData!现价
        End If
    Else
        '取时价药品售价
        gstrSQL = "select Decode(Nvl(零售价, 0), 0, Decode(Nvl(实际数量, 0), 0, 0, 实际金额 / 实际数量), 零售价) as 零售价 " & _
            " from 药品库存 where 性质=1 and  药品id=[1] and 库房id=[2] and nvl(批次,0)=[3]"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-零售价", lng药品id, lng库房id, lng批次)
        
        If rsData.EOF Then
            '时价药品零售价计算公式:采购价*(1+加成率)
            '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
            '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
            gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "零售价", lng药品id)
            dbl指导零售价 = rsData!指导零售价
            dbl差价让利比 = rsData!差价让利比
            
            Get售价 = 0
            dbl成本价 = Get成本价(lng药品id, lng库房id, lng批次)
            dbl加成率 = rsData!加成率 / 100
            dbl零售价 = dbl成本价 * (1 + dbl加成率)
            dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
            Get售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价)
        Else
            If rsData!零售价 = 0 Then
                gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
                Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "零售价", lng药品id)
                dbl指导零售价 = rsData!指导零售价
                dbl差价让利比 = rsData!差价让利比
                
                Get售价 = 0
                dbl成本价 = Get成本价(lng药品id, lng库房id, lng批次)
                dbl加成率 = rsData!加成率 / 100
                dbl零售价 = dbl成本价 * (1 + dbl加成率)
                dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
                Get售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价)
            Else
                Get售价 = rsData!零售价
            End If
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get时价零售价(ByVal lng药品id As Long, ByVal lng库房id As Long, ByVal lng批次 As Long, ByVal dbl比例系数 As Double) As Double
    '功能：获取时价药品当前药品的零售价
    '参数:药品id,库房id,批次
    '返回值：零售价
    Dim rsData As ADODB.Recordset
    Dim dbl零售价 As Double, dbl指导零售价 As Double, dbl差价让利比 As Double, dbl加成率 As Double
    Dim dbl成本价 As Double
    
    On Error GoTo errHandle
    gstrSQL = "select Decode(Nvl(零售价, 0), 0, Decode(Nvl(实际数量, 0), 0, 0, 实际金额 / 实际数量), 零售价) as 零售价 from 药品库存 where 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "零售价", lng药品id, lng库房id, lng批次)
    
    If rsData.EOF Then
        '时价药品零售价计算公式:采购价*(1+加成率)
        '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
        '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
        gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "零售价", lng药品id)
        dbl指导零售价 = rsData!指导零售价
        dbl差价让利比 = rsData!差价让利比
        
        Get时价零售价 = 0
        dbl成本价 = Get成本价(lng药品id, lng库房id, lng批次)
        dbl加成率 = rsData!加成率 / 100
        dbl零售价 = dbl成本价 * (1 + dbl加成率)
        dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
        Get时价零售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价) * dbl比例系数
    Else
        If rsData!零售价 = 0 Then
            gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "零售价", lng药品id)
            dbl指导零售价 = rsData!指导零售价
            dbl差价让利比 = rsData!差价让利比
            
            Get时价零售价 = 0
            dbl成本价 = Get成本价(lng药品id, lng库房id, lng批次)
            dbl加成率 = rsData!加成率 / 100
            dbl零售价 = dbl成本价 * (1 + dbl加成率)
            dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
            Get时价零售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价) * dbl比例系数
        Else
            Get时价零售价 = rsData!零售价 * dbl比例系数
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get零售价(ByVal bln是否时价 As Boolean, lng药品id As Long, ByVal lng库房id As Long, ByVal lng批次 As Long) As Double
    '功能：获取原始的售价单位售价，主要用于出库
    '参数: bln是否时价:false-定价,true-时价
    '返回值：最小单位的价格
    Dim rsData As ADODB.Recordset
    Dim dbl零售价 As Double, dbl指导零售价 As Double, dbl差价让利比 As Double, dbl加成率 As Double
    Dim dbl成本价 As Double
    
    On Error GoTo errHandle

    '取定价药品售价
    If bln是否时价 = False Then
        gstrSQL = "Select 现价 " & _
            " From 收费价目 A, 药品规格 B " & _
            " Where A.收费细目id = B.药品id And A.收费细目ID=[1] And Sysdate Between A.执行日期 And Nvl(A.终止日期,Sysdate) " & GetPriceClassString("A")
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "Get零售价-取定价药品售价", lng药品id)
        
        If Not rsData.EOF Then
            Get零售价 = rsData!现价
        End If
    Else
        '取时价药品售价
        gstrSQL = "select Decode(Nvl(零售价, 0), 0, Decode(Nvl(实际数量, 0), 0, 0, 实际金额 / 实际数量), 零售价) as 零售价 " & _
            " from 药品库存 where 性质=1 and  药品id=[1] and 库房id=[2] and nvl(批次,0)=[3]"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-零售价", lng药品id, lng库房id, lng批次)
        
        If rsData.EOF Then
            '时价药品零售价计算公式:采购价*(1+加成率)
            '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
            '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
            gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "零售价", lng药品id)
            dbl指导零售价 = rsData!指导零售价
            dbl差价让利比 = rsData!差价让利比
            
            Get零售价 = 0
            dbl成本价 = Get成本价(lng药品id, lng库房id, lng批次)
            dbl加成率 = rsData!加成率 / 100
            dbl零售价 = dbl成本价 * (1 + dbl加成率)
            dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
            Get零售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价)
        Else
            If rsData!零售价 = 0 Then
                gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1]"
                Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "零售价", lng药品id)
                dbl指导零售价 = rsData!指导零售价
                dbl差价让利比 = rsData!差价让利比
                
                Get零售价 = 0
                dbl成本价 = Get成本价(lng药品id, lng库房id, lng批次)
                dbl加成率 = rsData!加成率 / 100
                dbl零售价 = dbl成本价 * (1 + dbl加成率)
                dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
                Get零售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价)
            Else
                Get零售价 = rsData!零售价
            End If
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function CreateObject_LED(lngLEDModal As Long) As Boolean
    '创建LED显示对象
    
    Dim strsql As String
    Dim strObject As String

    On Error GoTo ErrHand
    
    '读取该LED显示接口的注册信息
    If grsLEDComponent.State = 0 Then
        strsql = "Select 部件类型,部件名,Nvl(启用,0) AS 启用 From 排队LED显示部件  "
        Set grsLEDComponent = zldatabase.OpenSQLRecord(strsql, "提取该LED显示接口的注册信息")
    End If
    grsLEDComponent.Filter = "部件类型=" & lngLEDModal
    If grsLEDComponent.RecordCount = 0 Then
        grsLEDComponent.Filter = 0
        MsgBox "该LED接口还未注册！ 序号=" & lngLEDModal, vbInformation, "排队叫号系统"
        Exit Function
    End If
    strObject = UCase(grsLEDComponent!部件名)
    grsLEDComponent.Filter = 0
    
    '检查该对象是否存在
    On Error Resume Next
    If Not gobjLEDShow Is Nothing Then
        CreateObject_LED = True
        Exit Function
    End If
    
    '去掉文件名后缀
    strObject = Mid(strObject, 1, Len(strObject) - 4)
    '创建对象
    Set gobjLEDShow = CreateObject(strObject & ".Cls" & Mid(strObject, 4))
    
    
    '调用初始化函数
    CreateObject_LED = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub SetGridFocus(ByVal objGrid As VSFlexGrid, ByVal blnGetFoucs As Boolean)
    With objGrid
        If blnGetFoucs Then
            .GridColorFixed = &H80000008
            .GridColor = &H80000008
            .ForeColorFixed = glngFixedForeColorByFocus
            .BackColorSel = glngRowByFocus
        Else
            .GridColorFixed = &H80000011
            .GridColor = &H80000011
            .ForeColorFixed = glngFixedForeColorNotFocus
            .BackColorSel = glngRowByNotFocus
        End If
    End With
End Sub
Public Function CheckIsCenter(ByVal lngStockid As Long) As Boolean
    '返回药房是否具有‘配制中心’性质
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From 部门性质说明 Where 工作性质 = '配制中心' And 部门id = [1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否具有配制中心性质", lngStockid)
    
    If Not rsTmp.EOF Then CheckIsCenter = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function Get现价(ByVal lng药品id As Long) As Double
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 现价 " & _
            " From 收费价目 A, 药品规格 B " & _
            " Where A.收费细目id = B.药品id And A.收费细目ID=[1] And Sysdate Between A.执行日期 And Nvl(A.终止日期,Sysdate) " & GetPriceClassString("A")
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "[提取该药品的零售单位价格]", lng药品id)
    
    If Not rstemp.EOF Then
        Get现价 = rstemp!现价
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetDefaultRecipeColor() As String
    GetDefaultRecipeColor = CStr(gconlng普通) & ";" & _
                    CStr(gconlng急诊) & ";" & _
                    CStr(gconlng儿科) & ";" & _
                    CStr(gconlng麻醉) & ";" & _
                    CStr(gconlng精一) & ";" & _
                    CStr(gconlng精二)

End Function
Public Sub DeptSendWork_CheckDrugstore(ByVal strPrivs As String, ByVal lngUserID As Long, ByVal strStateNo As String)
    '检测药房设置否(中药房、西药房及成药房)，在类模块中对应窗口打开前检查
    'strPrivs：权限；
    'lngUserID：当前用户ID；
    'strStateNo：当前系统站点编号；
    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    
    On Error GoTo errHandle
    If IsInString(strPrivs, "所有药房", ";") Then
        gstrSQL = "(Select Distinct 部门ID From 部门性质说明 Where 工作性质 Like '%药房' And 服务对象 IN (2,3))"
    Else
        gstrSQL = "(Select distinct A.部门ID From 部门人员 A,部门性质说明 B " & _
                 " Where A.人员ID=[1] And A.部门ID=B.部门ID And B.工作性质 Like '%药房' And B.服务对象 IN (2,3))"
    End If
    gstrSQL = " Select Distinct P.ID,P.名称 From 部门表 P " & _
             " Where (P.站点 = '" & strStateNo & "' Or P.站点 is Null) And P.ID In " & gstrSQL & _
             " And (P.撤档时间 Is Null Or P.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "检查药房设置", lngUserID)
    
    With rsData
        If .EOF Then
           If IsInString(strPrivs, "所有药房", ";") Then
               strMsg = "请初始化药房（部门管理）"
           Else
               strMsg = "你不是药房工作人员，不能操作本模块！"
           End If
           MsgBox strMsg, vbInformation, gstrSysName
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function DeptSendWork_GetDrugstore(ByVal strPrivs As String, ByVal lngUserID As Long, ByVal strStateNo As String) As ADODB.Recordset
    '取对应操作员允许操作的药房
    'strPrivs：权限；
    'lngUserID：当前用户ID；
    'strStateNo：当前系统站点编号；
    
    On Error GoTo errHandle
    If IsInString(strPrivs, "所有药房", ";") Then
        gstrSQL = "(Select Distinct 部门ID From 部门性质说明 Where 工作性质 Like '%药房' And 服务对象 IN (2,3))"
    Else
        gstrSQL = "(Select distinct A.部门ID From 部门人员 A,部门性质说明 B " & _
                 " Where A.人员ID=[1] And A.部门ID=B.部门ID And B.工作性质 Like '%药房' And B.服务对象 IN (2,3))"
    End If
    gstrSQL = " Select Distinct P.ID,P.名称 From 部门表 P " & _
             " Where (P.站点 = '" & strStateNo & "' Or P.站点 is Null) And P.ID In " & gstrSQL & _
             " And (P.撤档时间 Is Null Or P.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set DeptSendWork_GetDrugstore = zldatabase.OpenSQLRecord(gstrSQL, "取允许发药的药房", lngUserID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get给药途径() As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select 名称 as 用法 ,标本部位 As 分类 From 诊疗项目目录 Where 类别='E' And 操作类型='2'And (服务对象=2 Or 服务对象=3) " & _
            " And (撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or 撤档时间 Is Null) Order by 编码 "
    Set DeptSendWork_Get给药途径 = zldatabase.OpenSQLRecord(gstrSQL, "提取给药途径")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get自定义发药类型() As ADODB.Recordset
    On Error GoTo errHandle
    '提取发药类型
    gstrSQL = "Select 名称 From 发药类型 Order By 编码"
    Set DeptSendWork_Get自定义发药类型 = zldatabase.OpenSQLRecord(gstrSQL, "[提取发药类型]")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub MediWork_CheckInOutType()
    '检查药品入出类别
    Dim rsData As ADODB.Recordset
    Dim int入系数 As Integer, int出系数 As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select A.单据, A.类别id, B.ID, B.编码, B.名称, B.系数 " & _
        " From 药品单据性质 A, 药品入出类别 B " & _
        " Where A.类别id = B.Id " & _
        " Order By 单据"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "检查入出类别")
    
    With rsData
        If .EOF Then Exit Sub
        
        rsData.Filter = "单据=1"
        gInOutType.bln外购入库 = Not .EOF
        
        rsData.Filter = "单据=2"
        gInOutType.bln自制入库 = Not .EOF
        
        rsData.Filter = "单据=3"
        gInOutType.bln协药入库 = Not .EOF
        
        rsData.Filter = "单据=4"
        gInOutType.bln其它入库 = Not .EOF
        
        rsData.Filter = "单据=5"
        gInOutType.bln差价调整 = Not .EOF
        
        rsData.Filter = "单据=6"
        gInOutType.bln药品移库 = Not .EOF
        
        rsData.Filter = "单据=7"
        gInOutType.bln药品领用 = Not .EOF
        
        rsData.Filter = "单据=8"
        gInOutType.bln收费处方发药 = Not .EOF
        
        rsData.Filter = "单据=9"
        gInOutType.bln记帐单处方发药 = Not .EOF
        
        rsData.Filter = "单据=10"
        gInOutType.bln记帐表处方发药 = Not .EOF
        
        rsData.Filter = "单据=11"
        gInOutType.bln其他出库 = Not .EOF
        
        rsData.Filter = "单据=27"
        gInOutType.bln药品留存 = Not .EOF
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function DeptSendWork_CheckBill(ByVal IntOper As Integer, ByVal lng收发ID As Long, ByVal bln允许未审核处方发药 As Boolean) As Integer
    '--根据将要执行的操作，判断是否允许--
    '0-拒发;1-发药;2-退药
    '返回:
    '0-允许操作
    '1-已发药
    '2-已删除
    '3-未发药
    
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select A.NO,Nvl(B.记录状态,0) AS 审核标志,A.审核人,Decode(Nvl(A.摘要,'小宝'),'拒发',3,B.执行状态) 执行状态,A.发药方式 From 药品收发记录 A,住院费用记录 B " & _
             " Where A.费用ID=B.ID And A.ID=[1] "
    If IntOper = 2 Then
        gstrSQL = gstrSQL & " And A.审核人 IS Not Null"
    Else
        gstrSQL = gstrSQL & " And A.审核人 IS Null"
    End If
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "检查单据状态", lng收发ID)
    
    With rsData
        If .EOF Then
            DeptSendWork_CheckBill = 2
            MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If Not IsNull(!审核人) Then
            If IntOper <> 2 Then
                DeptSendWork_CheckBill = 1
                MsgBox "该处方[" & !NO & "]已被其它操作员发药，操作被迫中止！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If IntOper = 2 Then
                DeptSendWork_CheckBill = 3
                MsgBox "该处方[" & !NO & "]还未发药，操作被迫中止！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If IntOper = 1 Then
            If !执行状态 = 3 Then
                DeptSendWork_CheckBill = 2
                MsgBox "该处方[" & !NO & "]已拒发，操作被迫中止！", vbInformation, gstrSysName
                Exit Function
            End If
            
            If !审核标志 = 0 And bln允许未审核处方发药 = False Then
                DeptSendWork_CheckBill = 4
                MsgBox "该处方[" & !NO & "]还未审核，操作被迫中止！", vbInformation, gstrSysName
                Exit Function
            End If
            
            If Nvl(!发药方式, 0) = -1 Then
                DeptSendWork_CheckBill = 5
                MsgBox "该处方[" & !NO & "]已停止发药，操作被迫中止！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    DeptSendWork_CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MediWork_CheckStorageStock(ByVal lngStockid As Long, ByVal lngMediID As Long) As Boolean
    '检查指定药品是否设置存储库房
    'lngStockID：库房ID
    'lngMediID：药品ID
    
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 收费细目id From 收费执行科室 Where 执行科室id = [1] And 收费细目id = [2]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "检查药品存储库房", lngStockid, lngMediID)
    
    MediWork_CheckStorageStock = Not rsData.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get配药人(ByVal lng药房id As Long) As ADODB.Recordset
    '提取药房人员
    On Error GoTo errHandle
    gstrSQL = "Select Distinct A.简码||'-'||A.姓名 As 姓名,A.姓名 名称" & _
             " From 人员表 A,部门人员 B,部门性质说明 C,人员性质说明 D " & _
             " Where A.Id=B.人员id And B.部门id=C.部门Id And D.人员id=A.Id And D.人员性质 = '药房发药人' " & _
             " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) AND B.部门id=[1] " & _
             " ORDER BY 姓名 "

    Set DeptSendWork_Get配药人 = zldatabase.OpenSQLRecord(gstrSQL, "提取药房人员", lng药房id)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function DeptSendWork_Get核查人(ByVal lng药房id As Long) As ADODB.Recordset
    On Error GoTo errHandle
    '提取药房人员
    gstrSQL = "Select 简码||'-'||姓名 As 姓名,姓名 As 名称 From 人员表 Where Id In (Select 人员id from 部门人员 Where 部门id=[1]) " & _
             " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) " & _
             " ORDER BY 姓名 "

    Set DeptSendWork_Get核查人 = zldatabase.OpenSQLRecord(gstrSQL, "提取药房人员", lng药房id)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get发药单格式(ByVal strRPTCode As String) As ADODB.Recordset
    '获取报表格式名称
    '参数：strRPTCode-报表编码
    On Error GoTo errHandle
    gstrSQL = "Select 说明 As 格式 From zltools.zlRPTFMTs Where 报表id = (Select ID From zltools.zlReports Where 系统 = [1] And 编号 = [2]) Order By 序号"
    Set DeptSendWork_Get发药单格式 = zldatabase.OpenSQLRecord(gstrSQL, "取发药单格式", glngSys, strRPTCode)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get剂型(ByVal lng库房id As Long) As ADODB.Recordset
    '提取所有剂型
    On Error GoTo errHandle
    gstrSQL = "Select Distinct J.编码||'-'||J.名称 剂型" & _
         " From 诊疗执行科室 A,药品特性 B,药品剂型 J " & _
         " Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称" & _
         " And A.执行科室ID=[1]" & _
         " Order By j.编码 || '-' || j.名称 "
    Set DeptSendWork_Get剂型 = zldatabase.OpenSQLRecord(gstrSQL, "提取库房药品剂型", lng库房id)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function DeptSendWork_Get给药途径分类() As ADODB.Recordset
    '提取给药途径分类
    On Error GoTo errHandle
    gstrSQL = "Select Distinct 标本部位 As 分类 From 诊疗项目目录 Where 类别 = 'E' And 操作类型 = '2' And 标本部位 Is Not Null"
    Set DeptSendWork_Get给药途径分类 = zldatabase.OpenSQLRecord(gstrSQL, "取给药途径分类")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function IsInString(ByVal strTarget As String, ByVal strOrigin As String, Optional strSplit As String = "") As Boolean
    '某个字符串是否包含另一个字符串
    'strTarget：目标字符串
    'strOrigin：原字符串
    'strSplit：分隔符（不为空时为精确匹配）
    '在strTarget中是否包含strOrigin
    
    IsInString = InStrB(strSplit & strTarget & strSplit, strSplit & strOrigin & strSplit) > 0
End Function

Public Function MediWork_GetCheckStockRule(ByVal lng库房id As Long) As Integer
    '取出库检查规则
    Dim rsData As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取出库检查规则", lng库房id)

    If Not rsData.EOF Then
        MediWork_GetCheckStockRule = rsData!库存检查
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MediWork_GetMediRealAmount(ByVal lng库房id As Long, ByVal lng药品id As Long, ByVal lng批次 As Long) As Double
    '取药品实际库存
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(实际数量, 0) As 实际数量 " & _
            " From 药品库存 " & _
            " Where 性质 = 1 And 库房id = [1] And 药品id = [2] And Nvl(批次, 0) = [3] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取药品库存实际数量", lng库房id, lng药品id, lng批次)

    If Not rsData.EOF Then
        MediWork_GetMediRealAmount = rsData!实际数量
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function RecipeSendWork_GetDiagnosis(ByVal int门诊 As Integer, ByVal LngID As Long, Optional ByVal lng主页ID As Long, Optional ByVal int显示诊断类型 As Integer) As String
    '取病人诊断信息
    '1.门诊病人：根据医嘱ID来取诊断记录
    '2，3住院病人：根据病人ID、主页ID来取诊断记录
    'int显示诊断类型:1-只查询中药诊断;2-只查询西药诊断
    Dim rsData As ADODB.Recordset
    Dim strTmp, strFilter As String
    Dim strReturn As String
    Dim str记录日期, str诊断类型 As String
    Dim int诊断类型, n As Integer
    Dim str诊断信息, str诊断描述 As String
    
    '1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
    
    If LngID = 0 Then Exit Function
    On Error GoTo errHandle
    If int门诊 = 1 Then
        gstrSQL = "Select A.诊断描述,A.是否疑诊 From 病人诊断记录 A, 病人诊断医嘱 B Where A.ID = B.诊断id And B.医嘱id = [1] And 取消时间 Is Null "

        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "RecipeSendWork_GetDiagnosis", LngID)
        
        With rsData
            Do While Not .EOF
                If Nvl(!诊断描述, "") <> "" Then
                    strReturn = IIf(strReturn = "", "", strReturn & "|") & !诊断描述 & IIf(Nvl(rsData!是否疑诊, 0) = 1, "（？）", "")
                End If
                
                .MoveNext
            Loop
        End With
    ElseIf int门诊 = 2 Then
        '按出院，入院，门诊优选顺序返还诊断
        '返回值格式：诊断类型,诊断描述;诊断描述|诊断类型,诊断描述;诊断描述...
        gstrSQL = "Select 记录来源,诊断类型,诊断次序,诊断描述,是否疑诊,Mod(诊断类型,10) as 大类 From 病人诊断记录" & _
            " Where 病人ID=[1] And 主页ID=[2] And 诊断类型 IN(" & IIf(int显示诊断类型 = 1, "11,12,13", IIf(int显示诊断类型 = 2, "1,2,3", "1,2,3,11,12,13")) & ")" & _
            " Order by 记录来源,诊断类型,诊断次序"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "RecipeSendWork_GetDiagnosis", LngID, lng主页ID)
        
        '先按来源优先顺序过滤
        rsData.Filter = "记录来源=3" '首页整理
        If rsData.EOF Then rsData.Filter = "记录来源=2" '入院登记
        If rsData.EOF Then rsData.Filter = "记录来源=1" '病历
        If rsData.EOF Then rsData.Filter = "记录来源=4" '病案室录入
        
'        '住院再按类型优先顺序过滤
'        If Not rsData.EOF And int门诊 = 2 Then
'            gstrSQL = rsData.Filter
'            rsData.Filter = gstrSQL & " And 大类=3"
'            If rsData.EOF Then rsData.Filter = gstrSQL & " And 大类=2"
'            If rsData.EOF Then rsData.Filter = gstrSQL & " And 大类=1"
'        End If
        
        '住院再按类型优先顺序过滤
        strFilter = rsData.Filter
        For n = 3 To 1 Step -1
            str诊断描述 = ""
            rsData.Filter = strFilter & " And 大类=" & n
            Do While Not rsData.EOF
                Select Case rsData!诊断类型
                    Case 1
                        str诊断类型 = "西医门诊诊断"
                    Case 2
                        str诊断类型 = "西医入院诊断"
                    Case 3
                        str诊断类型 = "西医出院诊断"
                    Case 11
                        str诊断类型 = "中医门诊诊断"
                    Case 12
                        str诊断类型 = "中医入院诊断"
                    Case 13
                        str诊断类型 = "中医出院诊断"
                End Select
                
                If Not IsNull(rsData!诊断描述) Then
                    str诊断描述 = IIf(str诊断描述 = "", "", str诊断描述 & ";") & rsData!诊断描述 & IIf(Nvl(rsData!是否疑诊, 0) = 1, "（？）", "")
                End If
                
                rsData.MoveNext
                
                If rsData.EOF Then str诊断信息 = IIf(str诊断信息 = "", "", str诊断信息 & "|") & str诊断类型 & "," & str诊断描述
            Loop
        Next
        
        strReturn = str诊断信息
    Else
        '提取诊断
        '返回值格式：诊断类型,诊断描述;诊断描述|诊断类型,诊断描述;诊断描述...
        '按记录日期倒序排序，如果诊断类型，记录时间一致则合并诊断描述
        gstrSQL = "Select 诊断类型, 诊断描述, 记录日期,是否疑诊 " & _
            " From 病人诊断记录 " & _
            " Where 病人id = [1] And 主页id = [2] " & _
            " Order By 记录日期 Desc,诊断类型  Desc"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "RecipeSendWork_GetDiagnosis", LngID, lng主页ID)
        
        Do While Not rsData.EOF
            If str记录日期 & "," & int诊断类型 <> Format(rsData!记录日期, "YYYY-MM-DD HH:MM:SS") & "," & rsData!诊断类型 Then
                '诊断类型，记录日期与当前记录不相同时
                If str记录日期 <> "" Then
                    str诊断信息 = IIf(str诊断信息 = "", "", str诊断信息 & "|") & str诊断类型 & "," & str诊断描述
                End If
                
                Select Case rsData!诊断类型
                    Case 1
                        str诊断类型 = "西医门诊诊断"
                    Case 2
                        str诊断类型 = "西医入院诊断"
                    Case 3
                        str诊断类型 = "西医出院诊断"
                    Case 5
                        str诊断类型 = "院内感染"
                    Case 6
                        str诊断类型 = "病理诊断"
                    Case 7
                        str诊断类型 = "损伤中毒码"
                    Case 8
                        str诊断类型 = "术前诊断"
                    Case 9
                        str诊断类型 = "术后诊断"
                    Case 10
                        str诊断类型 = "并发症"
                    Case 11
                        str诊断类型 = "中医门诊诊断"
                    Case 12
                        str诊断类型 = "中医入院诊断"
                    Case 13
                        str诊断类型 = "中医出院诊断"
                    Case 21
                        str诊断类型 = "病原学诊断"
                    Case 22
                        str诊断类型 = "影像学诊断"
                End Select
                
                str记录日期 = Format(rsData!记录日期, "YYYY-MM-DD HH:MM:SS")
                int诊断类型 = rsData!诊断类型
                str诊断描述 = rsData!诊断描述 & IIf(Nvl(rsData!是否疑诊, 0) = 1, "（？）", "")
            Else
                '诊断类型，记录日期与当前记录相同时则合并诊断描述
                str诊断描述 = IIf(str诊断描述 = "", "", str诊断描述 & ";") & rsData!诊断描述 & IIf(Nvl(rsData!是否疑诊, 0) = 1, "（？）", "")
            End If
            
            rsData.MoveNext
            
            If rsData.EOF Then str诊断信息 = IIf(str诊断信息 = "", "", str诊断信息 & "|") & str诊断类型 & "," & str诊断描述
        Loop
        
        strReturn = str诊断信息
    End If
    
    RecipeSendWork_GetDiagnosis = strReturn
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get皮试结果(ByVal lng病人ID As Long, ByVal lng药名id As Long, ByVal dateCurrent As Date, ByVal date开嘱时间 As Date, Optional lng主页ID As Long) As String
    '取病人皮试结果，前提是传过来的药品属性是需要做皮试的药品
    '1、如果当前时间内（皮试结果有效天数设置）有皮试结果，就使用这个皮试结果。药房使用模块就显示为"阴性","阳性"或者"免试"。
    '2、如果当前时间内（皮试结果有效天数设置），没有皮试结果，就根据医嘱的开始执行时间和最后一次皮试结果登记时间比较，如果在皮试结果有效天数设置内，就使用这个皮试结果。药房使用模块就显示为“连续用药”。
    '3、如果1和2不成立，就显示“无皮试结果”
    Dim rsData As ADODB.Recordset
    
    If lng病人ID = 0 Then Exit Function
    
    On Error GoTo errHandle
    
'    gstrSQL = "Select 结果,记录时间 From 病人过敏记录 Where 病人id=[1] And 药物ID=[2] Order By 记录时间 Desc "
    
    gstrSQL = "Select Decode(结果, 1, '(+)', '(-)') As 结果, 记录时间 As 时间" & vbNewLine & _
        "From 病人过敏记录" & vbNewLine & _
        "Where 病人id = [1] And 药物id = [2]" & vbNewLine & _
        IIf(lng主页ID = 0, "", " And 主页id=[3] ") & _
        "Union All" & vbNewLine & _
        "Select '(免)' As 结果, a.开嘱时间 As 时间" & vbNewLine & _
        "From 病人医嘱记录 A" & vbNewLine & _
        "Where a.病人id = [1] And a.诊疗类别 = 'E' And 皮试结果='免试' And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From 诊疗项目目录 B, 诊疗用法用量 C" & vbNewLine & _
        "       Where b.Id = c.用法id And b.类别 = 'E' And b.操作类型 = '1' And b.Id = a.诊疗项目id And c.项目id = [2])" & vbNewLine & _
        IIf(lng主页ID = 0, "", " And A.主页id=[3] ") & _
        "Order By 时间 Desc"

    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取皮试结果", lng病人ID, lng药名id, lng主页ID)
    
    If rsData.RecordCount = 0 Then
        Get皮试结果 = "<无>"
        Exit Function
    ElseIf DateDiff("D", rsData!时间, dateCurrent) > gtype_UserSysParms.P70_过敏登记有效天数 Then
        '皮试时间距离当前时间超过期限天数
        If DateDiff("D", rsData!时间, date开嘱时间) > gtype_UserSysParms.P70_过敏登记有效天数 Then
            '开嘱时间期限天数前进行的皮试结果无效
            Get皮试结果 = "<无>"
            Exit Function
        Else
            '开嘱时间期限天数内进行的皮试有效
            Get皮试结果 = rsData!结果 & "<连>"
            Exit Function
        End If
    Else
        '皮试时间距离当前时间在有限天数内
        Get皮试结果 = rsData!结果
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function RecipeSendWork_Get医生() As ADODB.Recordset
    '取医生
    On Error GoTo errHandle
    gstrSQL = " Select Distinct A.简码||'-'||A.姓名 医生 From 人员表 A,人员性质说明 B" & _
             " Where B.人员性质='医生' And A.ID=B.人员ID" & _
             " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & _
             " Order by 医生"
    Set RecipeSendWork_Get医生 = zldatabase.OpenSQLRecord(gstrSQL, "取医生")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function RecipeSendWork_JudgeSign(ByVal Int单据 As Integer, ByVal strNo As String, Optional int可操作 As Integer, Optional ByVal lng收发ID As Long, Optional ByVal date时间 As Date) As Boolean
    Dim rsTmp As ADODB.Recordset
    
    '判断处方是否已进行了电子签名：返回真表示已有电子签名
    On Error GoTo errHandle
    If lng收发ID = 0 Then
        gstrSQL = "Select 1 From 药品签名明细 " & _
            " Where 收发id In (Select ID From 药品收发记录 Where 单据 = [1] And NO = [2] )  And Rownum = 1 "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "判断处方是否已进行了电子签名", Int单据, strNo)
    Else
        gstrSQL = "Select 1 From 药品签名明细 " & _
            " Where 收发id in (Select ID From 药品收发记录 Where Id=[3] And 单据 = [1] And NO = [2]) And  Rownum = 1 "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "判断处方是否已进行了电子签名", Int单据, strNo, lng收发ID)
    End If
    RecipeSendWork_JudgeSign = (rsTmp.RecordCount > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function RecipeSendWork_DispensingMedi(ByVal lng药房id As Long, bln是否配药确认 As Boolean) As Boolean
    '药房是否需要配药
    Dim rsData As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(配药,0) AS 配药,nvl(配药确认,0) as 配药确认,门诊 From 药房配药控制 Where 药房ID=[1] Order by 门诊"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取药房配药控制", lng药房id)
    
    '只要有一项表示需要经过配药过程的，标记为需要配药
    Do While Not rsData.EOF
        If rsData!配药 = 1 Then
            RecipeSendWork_DispensingMedi = True
        End If
        If rsData!门诊 = 1 Then
            If rsData!配药确认 = 1 Then
                bln是否配药确认 = True
            End If
        End If
        rsData.MoveNext
    Loop
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf列设置：列名，列宽，列对齐方式，固定列对齐方式（默认为居中对齐）
    
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColData(intCol) = lngColWidth
        
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub

Public Function TvwCheckNode(ByVal Node As Object, blnCheck As Boolean, Optional ByVal blnAutoExpand As Boolean = False)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        If blnAutoExpand = True Then Node.Expanded = blnCheck
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If blnAutoExpand = True Then Node.Expanded = blnCheck
            If Node.Children > 0 Then
                TvwCheckNode Node, blnCheck, blnAutoExpand
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function
Public Sub TvwSetParentNode(ByVal tvwObj As TreeView, ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = Node.FirstSibling.index
            Do While intIdx <> Node.LastSibling.index
                If tvwObj.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = tvwObj.Nodes(intIdx).Next.index
            Loop
            If intIdx = Node.LastSibling.index Then
                If tvwObj.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            TvwSetParentNode tvwObj, Node, blnCheck
        End If
    End If
End Sub

Public Function VerifySignatureRecored_bak(ByVal intTache As Integer, ByVal Int单据 As Integer, ByVal strNo As String, _
    ByVal lng药房id As Long, Optional ByVal LngID As Long, Optional ByVal date日期 As Date) As Boolean
    '电子签名
    'intTache:1-配药;2-发药
    Dim rsTmp As ADODB.Recordset
    Dim strSignSource As String
    Dim strDetail As String
    Dim strSign As String
    Dim lng证书ID As Long
    Dim strTimeStamp As String
    Dim strSignDate As String
    Dim intRule As Integer
    Dim lng签名id As Long
    
    '目前使用规则：
    intRule = 2
    On Error GoTo errHandle
    '获取签名源文
    gstrSQL = "Select A.ID, A.单据, A.NO, A.序号, A.库房id, A.入出类别id, A.对方部门id, A.入出系数, A.药品id, nvl(A.批次,0) 批次, " & _
        " A.填制人, To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') As 填制日期, A.配药人, To_Char(A.配药日期,'yyyy-MM-dd hh24:mi:ss') As 配药日期, A.审核人, To_Char(A.审核日期,'yyyy-MM-dd hh24:mi:ss') As 审核日期, " & _
        " A.费用id, A.单量, A.频次, A.用法, Nvl(B.签名ID, 0) As 签名ID " & _
        " From 药品收发记录 A, 药品签名明细 B,药品签名记录 C " & _
        " Where A.id=B.收发id and B.签名id=C.id and  单据 = [1] And No = [2] And 库房id + 0 = [3] "
    If LngID <> 0 Then
        gstrSQL = gstrSQL & " And a.id=[4] "
    Else
        If intTache = EsignTache.Dosage Then
            gstrSQL = gstrSQL & " And 配药日期=[4] And C.环节=1"
        Else
            gstrSQL = gstrSQL & " And 审核日期=[4] And C.环节<>1"
        End If
    End If
    
    gstrSQL = gstrSQL & " Order By 单据, NO, 序号,A.id "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "取单据信息", Int单据, strNo, lng药房id, IIf(LngID = 0, date日期, LngID))
    
    With rsTmp
        If Not .EOF Then
            strSignSource = !单据 & "," & !NO & "," & !库房id & "," & !入出类别id & "," & !对方部门id & "," & !入出系数
            
            If intTache = EsignTache.Dosage Then
                strSignSource = strSignSource & "," & !配药人 & "," & !配药日期
            Else
                strSignSource = strSignSource & "," & !审核人 & "," & !审核日期
            End If
        Else
            Exit Function
        End If
        
        strSignSource = strSignSource & "|"
        
        Do While Not .EOF
            lng签名id = !签名ID
            strDetail = IIf(strDetail = "", "", strDetail & ";") & !Id & "," & !序号 & "," & !药品ID & "," & Val(Nvl(!批次)) & "," & !费用ID & "," & !单量 & "," & !频次 & "," & !用法
            .MoveNext
        Loop
        
        strSignSource = strSignSource & strDetail
    End With
    
    '验证签名
    Call gobjESign.VerifySignature(strSignSource, lng签名id, 3)
    
    VerifySignatureRecored_bak = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function VerifySignatureRecoredGather(ByVal intTache As Integer, ByVal LngID As Long) As Boolean
    '验证电子签名：由于汇总发药签名时是签名一次，验证时只能汇总所有发药记录的信息来验证
    '需要注意保持和签名时的信息组成格式一致
    'intTache:1-配药;2-发药
    Dim rsTmp As ADODB.Recordset
    Dim strSignSource As String
    Dim strDetail As String
    Dim strSign As String
    Dim lng证书ID As Long
    Dim strTimeStamp As String
    Dim strSignDate As String
    Dim intRule As Integer
    Dim lng签名id As Long
    Dim Int单据 As Integer
    Dim strNo As String
    
    '目前使用规则：
    intRule = 2
    
    On Error GoTo errHandle
    
    '取签名ID
    gstrSQL = "Select b.签名id From 药品签名记录 A, 药品签名明细 B Where b.收发id = [1] And a.Id = b.签名id And a.环节 = 2 "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "取签名ID", LngID)
    If rsTmp.RecordCount > 0 Then
        lng签名id = rsTmp!签名ID
    Else
        Exit Function
    End If
        
    '获取签名源文；根据当前记录找到汇总发药时一并签名的所有记录
    gstrSQL = "Select A.ID, A.单据, A.NO, A.序号, A.库房id, A.入出类别id, A.对方部门id, A.入出系数, A.药品id, nvl(A.批次,0) 批次, " & _
        " A.填制人, To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') As 填制日期, A.配药人, To_Char(A.配药日期,'yyyy-MM-dd hh24:mi:ss') As 配药日期, A.审核人, To_Char(A.审核日期,'yyyy-MM-dd hh24:mi:ss') As 审核日期, " & _
        " A.费用id, A.单量, A.频次, A.用法, i.计算单位 " & _
        " From 药品收发记录 A, 诊疗项目目录 I, 药品规格 B " & _
        " Where a.药品id = b.药品id And i.Id = b.药名id And a.Id In (Select 收发id From 药品签名明细 Where 签名id = [1]) " & _
        " Order By 单据, NO, 序号 "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "取单据信息", lng签名id)
    
    With rsTmp
        Do While Not .EOF
            If Int单据 <> !单据 Or strNo <> !NO Then
                '单据信息与明细信息之间用|分隔
                If strDetail <> "" Then strSignSource = strSignSource & "|" & strDetail
                
                '不同单据之间用#分隔
                strSignSource = IIf(strSignSource = "", "", strSignSource & "#") & !单据 & "," & !NO & "," & !库房id & "," & !入出类别id & "," & !对方部门id & "," & !入出系数
                If intTache = EsignTache.send Or intTache = EsignTache.returnStep Then
                    strSignSource = strSignSource & "," & IIf(IsNull(!审核人), "", !审核人) & "," & IIf(IsNull(!审核日期), "", Format(!审核日期, "yyyy-MM-dd HH:mm:ss"))
                End If
                
                Int单据 = !单据
                strNo = !NO
                strDetail = ""
            End If
            
            '同一单据不同明细之间用;分隔
            strDetail = IIf(strDetail = "", "", strDetail & ";") & !Id & "," & !序号 & "," & !药品ID & "," & Val(Nvl(!批次)) & "," & !费用ID & "," & IIf(IsNull(!单量), "", FormatEx(!单量, 5) & Nvl(!计算单位)) & "," & IIf(IsNull(!频次), "", !频次) & "," & IIf(IsNull(!用法), "", !用法)
            
            .MoveNext
        Loop
    End With
    
    If strDetail <> "" Then strSignSource = strSignSource & "|" & strDetail

    '验证签名
    Call gobjESign.VerifySignature(strSignSource, lng签名id, 3)
    
    VerifySignatureRecoredGather = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetCheck库房(ByVal lng库房id As Long) As Integer
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1] "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "获取是否库存检查设置", lng库房id)
    If Not rstemp.EOF Then GetCheck库房 = Nvl(rstemp!库存检查, 0)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSignatureRecored(ByVal intTache As Integer, ByVal Int单据 As Integer, ByVal strNo As String, _
        ByVal lng药房id As Long, ByRef str签名记录 As String, Optional ByVal LngID As Long, _
        Optional ByVal date日期 As Date, Optional str操作人 As String, Optional lng现药房id As Long = 0, Optional blnCheck As Boolean) As Boolean
    '电子签名
    'intTache:1-配药;2-发药
    Dim rsTmp As ADODB.Recordset
    Dim strSignSource As String
    Dim strDetail As String
    Dim strSign As String
    Dim lng证书ID As Long
    Dim strTimeStamp As String
    Dim strTimeStampInfo As String
    Dim str收发ids As String
    Dim strSignDate As String
    Dim intRule As Integer
    
    '目前使用规则：
    intRule = 2
    
    gstrSQL = "Select ID, 单据, NO, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, nvl(批次,0) 批次, " & _
        " 填制人, To_Char(填制日期,'yyyy-MM-dd hh24:mi:ss') As 填制日期, 配药人, To_Char(配药日期,'yyyy-MM-dd hh24:mi:ss') As 配药日期, 审核人, To_Char(审核日期,'yyyy-MM-dd hh24:mi:ss') As 审核日期, " & _
        " 费用id, 单量, 频次, 用法 " & _
        " From 药品收发记录 " & _
        " Where  单据 = [1] And No = [2] And 库房id + 0 = [3] "
    If LngID <> 0 Then
        gstrSQL = gstrSQL & " And id=[4] "
    Else
        If intTache = EsignTache.Dosage Then
            gstrSQL = gstrSQL & " And Mod(记录状态,3)=1  And 审核人 Is Null "
        ElseIf intTache = EsignTache.send Then
            gstrSQL = gstrSQL & " And 审核人 Is Null  "
        ElseIf intTache = EsignTache.returnStep Then
            gstrSQL = gstrSQL & " And 审核日期=[4] "
        End If
            
    End If
    
    gstrSQL = gstrSQL & " Order By 单据, NO, 序号 "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "取单据信息", Int单据, strNo, lng药房id, IIf(LngID = 0, date日期, LngID))
    
    With rsTmp
        If Not .EOF Then
            strSignSource = !单据 & "," & !NO & "," & IIf(lng现药房id = 0, !库房id, lng现药房id) & "," & !入出类别id & "," & !对方部门id & "," & !入出系数
                
            If intTache = EsignTache.Dosage Then
                If str操作人 <> "" Then
                    strSignSource = strSignSource & "," & str操作人 & "," & Format(date日期, "yyyy-mm-dd hh:mm:ss")
                Else
                    strSignSource = strSignSource & "," & !配药人 & "," & !配药日期
                End If
            ElseIf intTache = EsignTache.send Then
                strSignSource = strSignSource & "," & str操作人 & "," & Format(date日期, "yyyy-mm-dd hh:mm:ss")
            ElseIf intTache = EsignTache.returnStep Then
                strSignSource = strSignSource & "," & "," & !审核人 & "," & !审核日期
            End If
        Else
            Exit Function
        End If
        
        strSignSource = strSignSource & "|"
        
        Do While Not .EOF
            str收发ids = IIf(str收发ids = "", "", str收发ids & ",") & !Id
            strDetail = IIf(strDetail = "", "", strDetail & ";") & !Id & "," & !序号 & "," & !药品ID & "," & Val(Nvl(!批次)) & "," & !费用ID & "," & !单量 & "," & !频次 & "," & !用法
            .MoveNext
        Loop
        
        strSignSource = strSignSource & strDetail
    End With
    
    strSign = gobjESign.Signature(strSignSource, gstrDbUser, lng证书ID, strTimeStamp, , strTimeStampInfo, blnCheck)
    If strSign <> "" Then
        If strTimeStamp <> "" Then
            strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
        Else
            strTimeStamp = "NULL"
        End If
        
        If strTimeStampInfo = "" Then strTimeStampInfo = "NULL"
        
        str签名记录 = intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & "," & strTimeStamp & ",'" & strTimeStampInfo & "'," & intTache & ",'" & str收发ids & "'"
        GetSignatureRecored = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetSignatureRecoredGather(ByVal intTache As Integer, ByVal rsData As ADODB.Recordset, ByVal lng库房id As Long, ByVal str配药人 As String, ByVal str审核人 As String, ByVal str审核日期 As String, ByRef str签名记录 As String, Optional blnCheck As Boolean) As Boolean
    '电子签名：用于汇总发药，每次发药操作算一次签名
    '直接从发药数据集组织数据，减少读取数据库操作
    'intTache:1-配药;2-发药
    Dim rsTmp As ADODB.Recordset
    Dim strSignSource As String
    Dim strDetail As String
    Dim strSign As String
    Dim lng证书ID As Long
    Dim strTimeStamp As String
    Dim strTimeStampInfo As String
    Dim str收发ids As String
    Dim strSignDate As String
    Dim intRule As Integer
    Dim Int单据 As Integer
    Dim strNo As String
    
    '目前使用规则：
    intRule = 2
    
    With rsData
'        .Filter = "执行状态=1"
        
        '排序方法不能变
        .Sort = "单据,NO,序号"
    
        Do While Not .EOF
            If Int单据 <> !单据 Or strNo <> !NO Then
                '单据信息与明细信息之间用|分隔
                If strDetail <> "" Then strSignSource = strSignSource & "|" & strDetail
                
                '不同单据之间用#分隔
                strSignSource = IIf(strSignSource = "", "", strSignSource & "#") & !单据 & "," & !NO & "," & lng库房id & "," & !入出类别id & "," & !领药部门ID & "," & !入出系数
                If intTache = EsignTache.send Or intTache = EsignTache.returnStep Then
                    strSignSource = strSignSource & "," & str审核人 & "," & str审核日期
                End If
                
                Int单据 = !单据
                strNo = !NO
                strDetail = ""
            End If
            
            str收发ids = IIf(str收发ids = "", "", str收发ids & ",") & !收发ID
            
            '同一单据不同明细之间用;分隔
            strDetail = IIf(strDetail = "", "", strDetail & ";") & !收发ID & "," & !序号 & "," & !药品ID & "," & Val(Nvl(!批次)) & "," & !费用ID & "," & !原始单量 & "," & !频次 & "," & !用法
            
            .MoveNext
        Loop
    End With
    
    If strDetail <> "" Then strSignSource = strSignSource & "|" & strDetail
    
    '获取签名信息
    strSign = gobjESign.Signature(strSignSource, gstrDbUser, lng证书ID, strTimeStamp, , strTimeStampInfo, blnCheck)
    If strSign <> "" Then
        If strTimeStamp <> "" Then
            strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
        Else
            strTimeStamp = "NULL"
        End If
        
        If strTimeStampInfo = "" Then strTimeStampInfo = "NULL"
        
        str签名记录 = intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & "," & strTimeStamp & ",'" & strTimeStampInfo & "'," & intTache & ",'" & str收发ids & "'"
        GetSignatureRecoredGather = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function






Public Function DelSignatureRecored_Check(ByVal intTache As Integer, ByVal Int单据 As Integer, ByVal strNo As String, ByVal lng药房id As Long, ByRef lng签名id As Long, Optional ByVal LngID As Long, Optional ByVal date日期 As Date) As Boolean
    'intRule:1-配药;2-发药
    Dim rsTmp As ADODB.Recordset
    Dim strSignSource As String
    Dim strDetail As String
    Dim strSign As String
    
    On Error GoTo errHandle
    '获取签名源文
    gstrSQL = "Select A.ID, A.单据, A.NO, A.序号, A.库房id, A.入出类别id, A.对方部门id, A.入出系数, A.药品id, nvl(A.批次,0) 批次, " & _
        " A.填制人, To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') As 填制日期, A.配药人, To_Char(A.配药日期,'yyyy-MM-dd hh24:mi:ss') As 配药日期, A.审核人, To_Char(A.审核日期,'yyyy-MM-dd hh24:mi:ss') As 审核日期, " & _
        " A.费用id, A.单量, A.频次, A.用法, Nvl(B.签名ID, 0) As 签名ID " & _
        " From 药品收发记录 A, 药品签名明细 B" & _
        " Where A.id=B.收发id(+) and 单据 = [1] And No = [2] And 库房id + 0 = [3] "
    If LngID <> 0 Then
        gstrSQL = gstrSQL & " And A.id=[4] "
    Else
        If intTache = EsignTache.Dosage Then
            gstrSQL = gstrSQL & " And Mod(记录状态,3)=1 "
        Else
            gstrSQL = gstrSQL & " And 审核日期=[4] "
        End If
            
    End If
    
    gstrSQL = gstrSQL & " Order By 单据, NO, 序号,Id "
    
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "取单据信息", Int单据, strNo, lng药房id, IIf(LngID = 0, date日期, LngID))
    
    With rsTmp
        If Not .EOF Then
            If CLng(!签名ID) = 0 Then
                '这种情况属于在业务发生时没有使用电子签名，而现在使用了，这种情况就不处理电子签名，而允许其他药房操作。
                DelSignatureRecored_Check = True
                Exit Function
            End If
            
            '检查USB-KEY
            If Not gobjESign.CheckCertificate(gstrDbUser) Then Exit Function
            
            lng签名id = CLng(!签名ID)
            strSignSource = !单据 & "," & !NO & "," & !库房id & "," & !入出类别id & "," & !对方部门id & "," & !入出系数 & "," & _
                !填制人 & "," & !填制日期 & "," & !配药人 & "," & !配药日期
            If intTache = EsignTache.send Then
                strSignSource = strSignSource & "," & !审核人 & "," & !审核日期
            End If
        Else
            Exit Function
        End If
        
        strSignSource = strSignSource & "|"
        
        Do While Not .EOF
            strDetail = IIf(strDetail = "", "", strDetail & ";") & !Id & "," & !序号 & "," & !药品ID & "," & Val(Nvl(!批次)) & "," & !费用ID & "," & !单量 & "," & !频次 & "," & !用法
            .MoveNext
        Loop
        
        strSignSource = strSignSource & strDetail
    End With
    
    '验证签名
    If Not gobjESign.VerifySignature(strSignSource, lng签名id, 3) Then Exit Function
    DelSignatureRecored_Check = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub CheckStopMedi(ByVal varInput As Variant, Optional ByRef Int退药 As Integer)
    '检查药品是否停用
    'varInput两种格式：传入单据信息（单据|No）;传入药品ID串（格式：药品ID1，药品ID2.....）
    'int退药:0-不是退药，1-退药，2-退药中有停用药品
    Dim rstemp As ADODB.Recordset
    Dim strMsg As String
    Dim Int单据 As Integer
    Dim strNo As String
    Dim n As Integer
    
    On Error GoTo errHandle
    If InStr(varInput, "|") > 0 Then
        Int单据 = Mid(varInput, 1, InStr(varInput, "|") - 1)
        strNo = Mid(varInput, InStr(varInput, "|") + 1)
        
        gstrSQL = "Select /*+rule*/ Distinct '(' || C.编码 || ')' || Nvl(B.名称, C.名称) As 药品信息 " & _
                " From 药品收发记录 A, 收费项目别名 B, 收费项目目录 C " & _
                " Where A.药品id = C.ID And A.药品id = B.收费细目id(+) And B.性质(+) = 3 " & _
                " And Nvl(C.撤档时间, To_Date('3000-01-01', 'yyyy-MM-dd')) <> To_Date('3000-01-01', 'yyyy-MM-dd') " & _
                " And A.单据 = [1] And A.NO = [2]"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "检查停用药品", Int单据, strNo)
    Else
        gstrSQL = "Select /*+ Rule*/ Distinct '(' || C.编码 || ')' || Nvl(B.名称, C.名称) As 药品信息 " & _
                " From Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) A, 收费项目别名 B, 收费项目目录 C " & _
                " Where A.Column_Value = C.ID  And A.Column_Value = B.收费细目id(+) And B.性质(+) = 3 " & _
                " And Nvl(C.撤档时间, To_Date('3000-01-01', 'yyyy-MM-dd')) <> To_Date('3000-01-01', 'yyyy-MM-dd') "
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "检查停用药品", varInput)
    End If
    
    With rstemp
        If Not .EOF Then
            For n = 1 To .RecordCount
                If n > 5 Then
                    strMsg = strMsg & vbCrLf & "还有其他" & .RecordCount - 5 & "个药品......"
                    Exit For
                End If
                strMsg = IIf(strMsg = "", "", strMsg & "," & vbCrLf) & !药品信息
                .MoveNext
            Next
            
            strMsg = "注意，以下药品已被停用：" & vbCrLf & strMsg
        End If
    End With
    
    If strMsg <> "" Then
        If Int退药 <> 0 Then
            MsgBox strMsg & vbCrLf & "停用的药品不允许退药。必须启用该药品，才可以进行退药操作", vbInformation, gstrSysName
            Int退药 = 2
        Else
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function Check库房分批(ByVal lng库房id As Long, ByVal lng药品id As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    Dim bln库存是否分批 As Boolean, bln分批 As Boolean, bln库房 As Boolean
    '分批返回true，不分批返回false
    On Error GoTo errHandle
    Check库房分批 = False
    
    '先判断是否是库房
    gstrSQL = "select 部门ID from 部门性质说明 where (工作性质 like '%药房' Or 工作性质 like '%制剂室') And 部门id=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "取部门性质", lng库房id)
    
    bln库房 = (rsCheck.EOF)
        
    '判断对应的药品目录中的分批属性
    gstrSQL = " Select Nvl(药库分批,0) 分批核算,nvl(药房分批,0) 药房分批核算 " & _
              " From 药品规格 Where 药品ID=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "取药品目录中的分批属性", lng药品id)
              
    If bln库房 Then
        Check库房分批 = (rsCheck!分批核算 = 1)
    Else
        Check库房分批 = (rsCheck!药房分批核算 = 1)
    End If

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckNumStock(ByVal objVSF As BillEdit, ByVal lng库房id As Long, ByVal lntCol药品id As Integer, _
    ByVal intCol批次 As Integer, ByVal intCol数量 As Integer, ByVal intCol比例系数 As Integer, _
    ByVal intMethod As Integer, Optional int入出业务 As Integer = 0, Optional int真实数量 As Integer = 0) As String
    '功能：审核出库类单据时，检查库存表实际数量是否足够
    '参数：objVSF-需要检查的表格;lng库房id；intcol批次-批次所在列；intCol数量-数量所在列；intCol比例系数-比例系数所在列
    '参数：int真实数量-真实数量所在列(用于审核，冲销)；intMethod，1-正常审核，2-冲销，3-退库审核
    '参数：int入出业务，0-入库；1-出库
    '返回值：哪行具体的药品名称，为空-检查通过，数量充足；不为空-检查未通过，数量不充足
    Dim objCol As Collection         '已使用的数量集合
    Dim dblNum As Double
    Dim varNum As Variant
    Dim varTemp As Variant
    Dim strTemp As String
    Dim lng药品id As Long
    Dim lng批次 As Long
    Dim rsData As ADODB.Recordset
    Dim strKey As String
    Dim vardrug As Variant
    Dim lngRow As Long
    Dim strArray As String
    
    On Error GoTo errHandle
    
    '先组合表格中数量，组合数量主要是考虑不分批的情况
    Set objCol = New Collection
    With objVSF
        If .rows < 2 Then Exit Function
        For lngRow = 1 To .rows - 1
            dblNum = 0
            If .TextMatrix(lngRow, lntCol药品id) <> "" Then
                For Each vardrug In objCol
                    If vardrug(0) = .TextMatrix(lngRow, lntCol药品id) & "," & Val(.TextMatrix(lngRow, intCol批次)) Then
                        dblNum = vardrug(1)
                        objCol.Remove vardrug(0)
                        Exit For
                    End If
                Next
                strKey = .TextMatrix(lngRow, lntCol药品id) & "," & Val(.TextMatrix(lngRow, intCol批次))
                
                '如果界面数量是小数，则按原始数据库的数量来计算
                If Fix(Val(.TextMatrix(lngRow, intCol数量))) <> Val(.TextMatrix(lngRow, intCol数量)) And int真实数量 <> 0 Then
                    strArray = dblNum + Val(.TextMatrix(lngRow, int真实数量))
                Else
                    strArray = dblNum + (Val(.TextMatrix(lngRow, intCol数量)) * Val(.TextMatrix(lngRow, intCol比例系数)))
                End If
                
                objCol.Add Array(strKey, strArray), strKey
            End If
        Next
    End With
    
    For Each varNum In objCol
        strTemp = varNum(0)  '格式是药品id,批次
        dblNum = varNum(1)
        varTemp = Split(strTemp, ",")
        If int入出业务 = 0 Then '入库
            If intMethod = 1 Then '正常审核
                If dblNum < 0 Then
                    '负数入库，需要减库存，所以需要判断库存是否充足
                    dblNum = Abs(dblNum)
                Else
                    '正数入库，不见库存，所以不检查
                    dblNum = 0
                End If
            ElseIf intMethod = 2 Then
                '冲销
                If dblNum < 0 Then
                    dblNum = 0
                Else
                    dblNum = dblNum
                End If
            ElseIf intMethod = 3 Then
                '退库审核，退库必须录入正数
                dblNum = dblNum
            End If
        Else    '出库
            If intMethod = 1 Then '正常审核
                If dblNum < 0 Then
                    '负数入库，需要减库存，所以需要判断库存是否充足
                    dblNum = 0
                Else
                    '正数入库，不见库存，所以不检查
                    dblNum = dblNum
                End If
            ElseIf intMethod = 2 Then
                '冲销
                If dblNum < 0 Then
                    dblNum = Abs(dblNum)
                Else
                    dblNum = 0
                End If
            End If
        End If
        
        '只有有数量才判断
        If dblNum > 0 Then
            lng药品id = varTemp(0)
            lng批次 = varTemp(1)
            If Check库房分批(lng库房id, lng药品id) = False Then
                lng批次 = 0
            End If
            
            gstrSQL = "Select (a.实际数量 - [1]) As 剩余数量, b.名称, b.编码" & vbNewLine & _
                        "From 药品库存 A, 收费项目目录 B" & vbNewLine & _
                        "Where a.药品id = b.Id And a.药品id = [2] And a.库房id = [3] And Nvl(a.批次, 0) = [4] And b.类别 In ('5', '6', '7') And a.性质 = 1"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "库存检查", dblNum, lng药品id, lng库房id, lng批次)
            
            If rsData.RecordCount = 0 Then
                gstrSQL = "select 编码,名称 from 收费项目目录 where id=[1]"
                Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "库存检查", lng药品id)
                CheckNumStock = "[" & rsData!编码 & "]" & rsData!名称
                Exit Function
            Else
                If rsData!剩余数量 >= 0 Then
                    CheckNumStock = ""
                Else
                    CheckNumStock = "[" & rsData!编码 & "]" & rsData!名称
                    Exit Function
                End If
            End If
        End If
    Next
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckUsableNum( _
    ByVal lng库房id As Long, _
    ByVal lng药品id As Long, _
    ByVal lng批次 As Long, _
    ByVal dbl填写数量 As Double, _
    ByVal dbl换算系数 As Double, _
    ByVal strNo As String, _
    ByVal Int单据 As Integer, _
    ByVal int库存检查 As Integer, _
    ByVal int数量精度 As Integer, _
    Optional int序号 As Integer, _
    Optional dblSum As Double) As Boolean
    '界面填写数量时用来检查可用数量是否足够，包括新增/修改，冲销等情况
    '返回值 true-通过检查，false-没有通过检查
    '入参：dbl填写数量是界面单位数量
    '      strNo="", 空-填单 非空-修改，修改时需要排除当前单据数量
    '      dblSum 界面该药品总填写数量，适用于冲销/申请冲销时
    '1.批次大于0是按批次检查，批次=0则是表示整体库存检查；修改状态时要考虑原单据数量；分批的要考虑可能被其他未进行批次分解的业务占用的数量
    '2.如果不需要检查库存的就不用调该函数，如出库冲销
    '3.申领/移库单据冲销时特殊处理:
    '根据序号取原入库的批次，注意要传原单据入库房(冲销时为出库房)，因暂时不支持对冲销申请的修改，所以不考虑已有单据的情况，要从界面传入总数量
    '4.提醒或禁止时根据分批还是总量不足有所不同
    Dim dblNum As Double
    Dim rsData As ADODB.Recordset
    Dim dblCheck As Boolean
    Dim bln分批不足 As Boolean
    Dim bln总量不足 As Boolean
    Dim strSqlStock As String, strSqlStockBatch As String  '库存数量，总数量和分批数量
    Dim strSqlSum As String, strSqlSumBatch As String      '库存合并未审核的数量，总数量和分批数量
    Dim lng出库批次 As Long
    Dim blnNewNo As Boolean '是否新增单据
    Dim dbl总填写数量 As Double
    
    On Error GoTo errHandle
    
    If int库存检查 = 0 Then CheckUsableNum = True: Exit Function

    If Int单据 = 6 And int序号 > 0 Then
        blnNewNo = True
        
        '取原入库的那笔的批次
        gstrSQL = "Select Nvl(批次, 0) 批次 From 药品收发记录 Where  " & _
            " 库房id=[1] And 单据 = [2] And NO = [3] And 序号 = [4] And 药品id = [5] And 入出系数 = 1"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取入库批次", lng库房id, Int单据, strNo, int序号 + 1, lng药品id)
        
        If rsData.RecordCount = 0 Then Exit Function
        
        lng出库批次 = rsData!批次
        
        If lng出库批次 = 0 Then
            '出库批次为不分批，按界面总数量
            dbl总填写数量 = dblSum
        Else
            '出库批次为分批，按界面该批次的填写数量
            dbl总填写数量 = dbl填写数量
        End If
    Else
        blnNewNo = (strNo = "")
        lng出库批次 = lng批次
        dbl总填写数量 = dbl填写数量
    End If
        
    strSqlStock = "Select Sum(Nvl(可用数量, 0)) As 可用数量 From 药品库存 Where 性质=1 And 库房id = [1] And 药品id = [2]"
    strSqlStockBatch = "Select Sum(Nvl(可用数量, 0)) As 可用数量 From 药品库存 Where 性质=1 And 库房id = [1] And 药品id = [2] And nvl(批次,0) = [3] "
    strSqlSum = "Select Sum(可用数量) As 可用数量" & vbNewLine & _
                " From (Select Nvl(可用数量, 0) As 可用数量" & vbNewLine & _
                "       From 药品库存" & vbNewLine & _
                "       Where 性质=1 And 库房id = [1] And 药品id = [2] " & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Abs(a.实际数量 * Nvl(a.付数, 1)) As 可用数量" & vbNewLine & _
                "       From 药品收发记录 A" & vbNewLine & _
                "       Where a.审核日期 Is Null And a.库房id = [1] And a.药品id + 0 = [2]  And a.No = [4] And a.单据 = [5])"
    strSqlSumBatch = "Select Sum(可用数量) As 可用数量" & vbNewLine & _
                    " From (Select Nvl(可用数量, 0) As 可用数量" & vbNewLine & _
                    "       From 药品库存" & vbNewLine & _
                    "       Where 性质=1 And 库房id = [1] And 药品id = [2]  And nvl(批次,0) = [3] " & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Abs(a.实际数量 * Nvl(a.付数, 1)) As 可用数量" & vbNewLine & _
                    "       From 药品收发记录 A" & vbNewLine & _
                    "       Where a.审核日期 Is Null And a.库房id = [1] And a.药品id + 0 = [2]  And a.No = [4] And a.单据 = [5]  And nvl(批次,0) = [3] )"
    
    If lng批次 = 0 Then
        '1.不分批的情况
        If blnNewNo = True Then
            '1.1如果是单据新增状态，直接看库存总可用数量是否足够
            gstrSQL = strSqlStock
        Else
            '1.2如果是单据修改状态，要合并原单据数量
            gstrSQL = strSqlSum
        End If
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "可用数量", lng库房id, lng药品id, lng出库批次, strNo, Int单据)
        
        If Nvl(rsData.Fields(0), 0) > 0 Then
            dblNum = zlStr.FormatEx(rsData.Fields(0) / dbl换算系数, int数量精度, True, False)
        End If
        
        If dblNum < dbl总填写数量 Then
            dblCheck = True
            bln分批不足 = True
        End If
    Else
        '2.分批的情况
        If blnNewNo = True Then
            '2.1如果是单据新增状态，直接看库存总可用数量是否足够
            gstrSQL = strSqlStockBatch
        Else
            '2.2如果是单据修改状态，要合并原单据数量
            gstrSQL = strSqlSumBatch
        End If
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "可用数量", lng库房id, lng药品id, lng出库批次, strNo, Int单据)

        If Nvl(rsData.Fields(0), 0) > 0 Then
            dblNum = zlStr.FormatEx(rsData.Fields(0) / dbl换算系数, int数量精度, True, False)
        End If
        
        If dblNum < dbl总填写数量 Then
            '2.2.1分批不够
            dblCheck = True
            bln分批不足 = True
        End If
    End If
        
    '库存不够时提醒或禁止
    If dblCheck = True Then
        gstrSQL = "select 编码,名称 from 收费项目目录 where id=[1]"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "库存检查", lng药品id)
                    
        Select Case int库存检查
        Case 1  '提示
            If Int单据 = 2 Then '自制入库
                If bln总量不足 = True Then
                    If MsgBox("组成药品【[" & rsData!编码 & "]" & rsData!名称 & "】的可用库存不足，可能被其他未审核单据占用，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf bln分批不足 = True Then
                    If MsgBox("组成药品【[" & rsData!编码 & "]" & rsData!名称 & "】出库数量大于了可用库存" & dblNum & "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            Else
                If bln总量不足 = True Then
                    If MsgBox("【[" & rsData!编码 & "]" & rsData!名称 & "】的可用库存不足，可能被其他未审核单据占用，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf bln分批不足 = True Then
                    If MsgBox("【[" & rsData!编码 & "]" & rsData!名称 & "】出库数量大于了可用库存" & dblNum & "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        Case 2  '禁止
            If Int单据 = 2 Then '自制入库
                If bln总量不足 = True Then
                    MsgBox "组成药品【[" & rsData!编码 & "]" & rsData!名称 & "】的可用库存不足，可能被其他未审核单据占用，不能出库！", vbInformation, gstrSysName
                ElseIf bln分批不足 = True Then
                    MsgBox "组成药品【[" & rsData!编码 & "]" & rsData!名称 & "】出库数量大于了可用库存" & dblNum & "，不能出库！", vbInformation, gstrSysName
                End If
            Else
                If bln总量不足 = True Then
                    MsgBox "【[" & rsData!编码 & "]" & rsData!名称 & "】的可用库存不足，可能被其他未审核单据占用，不能出库！", vbInformation, gstrSysName
                ElseIf bln分批不足 = True Then
                    MsgBox "【[" & rsData!编码 & "]" & rsData!名称 & "】出库数量大于了可用库存" & dblNum & "，不能出库！", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        End Select
    End If
    CheckUsableNum = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get分批属性(ByVal lng库房id As Long, ByVal lng药品id As Long) As Integer
    '返回指定库房，指定药品的分批属性
    '返回：0-不分批，1-分批
    Dim rsCheck As New ADODB.Recordset
    Dim int分批 As Integer
    Dim bln药房 As Boolean
    Dim strsql As String
        
    On Error GoTo errHandle
    
    '判断是否是药房或制剂室
    strsql = "select 部门ID from 部门性质说明 where (工作性质 like '%药房' Or 工作性质 like '%制剂室') And 部门id=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(strsql, "Get分批属性", lng库房id)

    bln药房 = (Not rsCheck.EOF)
        
    '判断对应的药品目录中的分批属性
    strsql = " Select Nvl(药库分批,0) As 药库分批,nvl(药房分批,0) As 药房分批 " & _
              " From 药品规格 Where 药品ID=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(strsql, "Get分批属性", lng药品id)
              
    If bln药房 Then
        int分批 = rsCheck!药房分批
    Else
        int分批 = rsCheck!药库分批
    End If
    
    Get分批属性 = int分批
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function CheckStrickUsable(ByVal Int单据 As Integer, ByVal lng库房id As Long, _
        ByVal lng药品id As Long, ByVal str药品名称 As String, _
        ByVal lng批次 As Long, ByVal dbl冲销数量 As Double, ByVal int库存检查 As Integer, _
        Optional ByVal strNo As String = "", Optional ByVal int序号 As Integer = 0) As Boolean
    '冲销单据时检查：原单据入库库房是否可用数量足够（可用数量等于或小于实际数量），实际冲销数量不能大于可用数量
    '对于移库单据、他入库单，需要取原单据入库那笔的批次，再根据批次来取可用数量；
    '对于自制入库、协定入库单据，由于是全部冲销，可以根据单据号，序号来取冲销数量，再来和库存可用数量比较
    '其他单据可直接根据批次取库存可用数量
    'int库存检查：表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
    '只有冲销时是出库类型（原单据是入库类型）的要做此检查：外国入库、自制入库（原单据入的那笔）、协定入库（原单据入的那笔）、其他入库、移库（原单据入的那笔）
    
    Dim rstemp As ADODB.Recordset
    Dim lng入库批次 As Long
    Dim dbl可用数量 As Double
    
    On Error GoTo errHandle
    
    If int库存检查 = 0 Then CheckStrickUsable = True: Exit Function
    
    If Int单据 = 2 Or Int单据 = 3 Then  '自制入库、协定入库单据
        If strNo = "" Or int序号 = 0 Then Exit Function
        gstrSQL = "Select 1 From 药品收发记录 A, 药品库存 B " & _
            " Where A.单据 = [1] And A.NO = [2] And A.序号 = [3] And A.记录状态 = 1 And A.入出系数 = 1 And B.性质 = 1 And A.库房id = B.库房id And A.药品id = B.药品id And " & _
            " Nvl(A.批次, 0) = Nvl(B.批次, 0) And A.实际数量 > B.可用数量 And Rownum = 1"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "检查可用数量", Int单据, strNo, int序号)
        
        '按正常流程进行提示或禁止
        If rstemp.RecordCount > 0 Then
            Select Case int库存检查
            Case 1  '提示
                If MsgBox(str药品名称 & "的库存不足，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Case 2  '禁止
                MsgBox str药品名称 & "的库存不足！", vbInformation, gstrSysName
                Exit Function
            End Select
        End If
    Else
        If Int单据 = 6 Or Int单据 = 4 Then   '移库单，其他入库单
            If strNo = "" Or int序号 = 0 Then Exit Function
            
            gstrSQL = "Select Nvl(批次, 0) 批次 From 药品收发记录 Where 单据 = [1] And NO = [2] And 序号 = [3] And 药品id = [4] And 入出系数 = 1"
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "取入库批次", Int单据, strNo, int序号, lng药品id)
            
            If rstemp.RecordCount = 0 Then Exit Function
            
            lng入库批次 = rstemp!批次
        Else
            '其他单据根据传入的批次来取库存可用数量
            lng入库批次 = lng批次
        End If
        
        gstrSQL = "Select Nvl(可用数量, 0) 可用数量 From 药品库存 Where 性质 = 1 And 库房id = [1] And 药品id = [2] And Nvl(批次, 0) = [3] "
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "取可用数量", lng库房id, lng药品id, lng入库批次)
        
        If rstemp.RecordCount > 0 Then
            dbl可用数量 = rstemp!可用数量
        End If
        
        '按正常流程进行提示或禁止
        If dbl可用数量 < Abs(dbl冲销数量) Then
            Select Case int库存检查
            Case 1  '提示
                If MsgBox(str药品名称 & "的库存不足，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Case 2  '禁止
                MsgBox str药品名称 & "的库存不足！", vbInformation, gstrSysName
                Exit Function
            End Select
        End If
    End If
    
    CheckStrickUsable = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub LoadBillControl()
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(时间限制, 0) 时间限制, Nvl(他人单据, 0) 他人单据, Nvl(金额上限, 0) 金额上限 From 单据操作控制 Where 人员id = [1] And 单据 = 9"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "单据操作控制", glngUserId)
    
    If Not rsTmp.EOF Then
        gtype_myBillControl.bln是否控制 = True
        gtype_myBillControl.int时间限制 = rsTmp!时间限制
        gtype_myBillControl.bln他人单据 = (rsTmp!他人单据 = 1)
        gtype_myBillControl.dbl金额上限 = rsTmp!金额上限
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckBillControl(ByVal IntOper As Integer, ByVal IntBillStyle As Integer, ByVal strNo As String, ByVal dblMoney As Double) As Boolean
    '--根据单据操作控制表，检查当前操作员是否允许操作单据
    'IntOper:1-配药;2-取消配药;3-发药;4-退药
    Dim rstemp As New ADODB.Recordset
    Dim bln是否初次发药 As Boolean
    
    On Error GoTo errHandle
    If gtype_myBillControl.bln是否控制 = False Then
        CheckBillControl = True
        Exit Function
    End If
    
    
    '检查时间限制
    If gtype_myBillControl.int时间限制 > 0 Then
        If IntOper <> 4 Then
            gstrSQL = "Select Distinct 填制日期 From 药品收发记录 Where 单据 = [1] And NO = [2] And Mod(记录状态, 3) = 1 And 记录状态 <> 1 And 审核人 Is Null"
        Else
            gstrSQL = "Select Distinct 填制日期 From 药品收发记录 Where 单据 = [1] And NO = [2] And Mod(记录状态, 3) = 1 And 审核人 Is Not Null"
        End If
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "检查单据操作控制", IntBillStyle, strNo)
         
        If Not rstemp.EOF Then
            If DateDiff("d", Format(rstemp!填制日期, "yyyy-mm-dd hh:mm:ss"), Sys.Currentdate) > gtype_myBillControl.int时间限制 Then
                MsgBox "处方[" & strNo & "]超过允许的最大操作时限，不能进行操作。"
                Exit Function
            End If
        Else
            bln是否初次发药 = True
        End If
    End If
    
    '检查是否允许操作他人单据
    If gtype_myBillControl.bln他人单据 Then
        If IntOper <> 4 Then
            gstrSQL = "Select 审核人 From 药品收发记录 Where 单据 = [1] And NO = [2] And Mod(记录状态, 3) = 2 And 审核人 Is Not Null Order By 审核日期 Desc"
        Else
            gstrSQL = "Select 审核人 From 药品收发记录 Where 单据 = [1] And NO = [2] And Mod(记录状态, 3) = 1 And 审核人 Is Not Null Order By 审核日期 Desc"
        End If
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "检查单据操作控制", IntBillStyle, strNo)
         
        If Not rstemp.EOF Then
            If rstemp!审核人 <> gstrUserName Then
                MsgBox "处方[" & strNo & "]上次操作人不是当前操作员，不能进行操作。"
                Exit Function
            End If
        End If
    End If
    
    '检查金额上限
    If gtype_myBillControl.dbl金额上限 > 0 And bln是否初次发药 = False Then
        If gtype_myBillControl.dbl金额上限 < dblMoney Then
            MsgBox "处方[" & strNo & "]金额超过允许操作的最大金额，不能进行操作。"
            Exit Function
        End If
    End If
    
    CheckBillControl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPrice(ByVal lngBillId As Long, ByRef strMsg As String) As Boolean
    Dim rstemp As New ADODB.Recordset
    '判断售价是否是当前最新售价
    
    CheckPrice = True       '统一由Oracle存储过程检查
    Exit Function
    
    '取原始价格和现价
    On Error GoTo errHandle
    gstrSQL = _
        "Select Round(a.原价, b.精度) 原价, Round(a.现价, b.精度) 现价, a.是否变价" & vbNewLine & _
        "From (Select Nvl(a.零售价, 0) 原价, b.现价, Nvl(c.是否变价, 0) 是否变价" & vbNewLine & _
        "       From 药品收发记录 A, 收费价目 B, 收费项目目录 C, 门诊费用记录 D" & vbNewLine & _
        "       Where a.药品id = b.收费细目id And a.药品id = c.Id And a.费用id = d.Id And" & vbNewLine & _
        "             (Sysdate Between b.执行日期 And b.终止日期 Or Sysdate >= b.执行日期 And b.终止日期 Is Null) And a.Id = [1]" & _
        GetPriceClassString("B") & ") A, 药品卫材精度 B" & vbNewLine & _
        "Where b.性质 = 0 And b.类别 = 1 And b.内容 = 2 And b.单位 = 1" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select Round(a.零售价, Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))) 原价," & vbNewLine & _
        "       Round(b.现价, Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))) 现价, Nvl(c.是否变价, 0) 是否变价" & vbNewLine & _
        "From 药品收发记录 A, 收费价目 B, 收费项目目录 C, 住院费用记录 D" & vbNewLine & _
        "Where a.药品id = b.收费细目id And a.药品id = c.Id And a.费用id = d.Id And" & vbNewLine & _
        "      (Sysdate Between b.执行日期 And b.终止日期 Or Sysdate >= b.执行日期 And b.终止日期 Is Null) And a.Id = [1] " & GetPriceClassString("B")
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "[取原始价格和最新价格]", lngBillId)
    
    If rstemp.RecordCount = 0 Then
        strMsg = "药品卫材精度表无数据！"
        CheckPrice = True
        Exit Function
    End If
    
    '时价药品不处理
    If rstemp!是否变价 = 1 Then
        CheckPrice = True
        Exit Function
    End If
    
    '比较价格
    If rstemp!原价 <> rstemp!现价 Then
        strMsg = "原价为" & rstemp!原价 & ",现价为" & rstemp!现价 & "。" & vbCrLf & Space(4) & "退药将产生调价退药明细记录，是否继续退药? "
        CheckPrice = False
        Exit Function
    End If
    
    CheckPrice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'取系统参数值
Public Sub GetSysParms()
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    
    gtype_UserSysParms.P6_未审核记帐处方发药 = zldatabase.GetPara(6, glngSys, , 0)
    gtype_UserSysParms.P9_费用金额保留位数 = zldatabase.GetPara(9, glngSys, , 0)
    gtype_UserSysParms.P15_门诊收费与发药分离 = zldatabase.GetPara(15, glngSys, , 0)
    gtype_UserSysParms.P16_住院记帐与发药分离 = zldatabase.GetPara(16, glngSys, , 0)
    gtype_UserSysParms.P23_已结帐单据操作 = zldatabase.GetPara(23, glngSys, , 0)
    gtype_UserSysParms.P25_使用电子签名 = zldatabase.GetPara(25, glngSys, , 0)
    gtype_UserSysParms.P26_电子签名场合 = zldatabase.GetPara(26, glngSys, , 0)
    gtype_UserSysParms.P28_门诊病人消费时需要刷卡验证 = zldatabase.GetPara(28, glngSys, , "1|0")
    gtype_UserSysParms.P29_指导批发价定价单位 = zldatabase.GetPara(29, glngSys, , 0)
    gtype_UserSysParms.P44_输入匹配 = zldatabase.GetPara(44, glngSys, , 0)
    gtype_UserSysParms.P54_时价药品以加价率入库 = zldatabase.GetPara(54, glngSys, , 0)
    gtype_UserSysParms.P64_审核限制 = zldatabase.GetPara(64, glngSys, , 0)
    gtype_UserSysParms.P68_门诊药嘱先作废后退药 = zldatabase.GetPara(68, glngSys, , 0)
    gtype_UserSysParms.P70_过敏登记有效天数 = zldatabase.GetPara(70, glngSys, , 0)
    gtype_UserSysParms.P75_外购入库需要核查 = zldatabase.GetPara(75, glngSys, , 0)
    gtype_UserSysParms.P76_时价药品直接确定售价 = zldatabase.GetPara(76, glngSys, , 0)
    gtype_UserSysParms.P85_药房查看单据成本价 = zldatabase.GetPara(85, glngSys, , 0)
    gtype_UserSysParms.P96_药品填单下可用库存 = zldatabase.GetPara(96, glngSys, , 0)
    gtype_UserSysParms.P98_记帐报警包含划价费用 = zldatabase.GetPara(98, glngSys, , 0)
    gtype_UserSysParms.P126_时价药品售价加成方式 = zldatabase.GetPara(126, glngSys, , 0)
    gtype_UserSysParms.P148_未收费处方发药 = zldatabase.GetPara(148, glngSys, , 0)
    gtype_UserSysParms.P149_效期显示方式 = zldatabase.GetPara(149, glngSys, , 0)
    gtype_UserSysParms.P150_药品出库优先算法 = zldatabase.GetPara(150, glngSys, , 0)
    gtype_UserSysParms.P153_配置中心 = zldatabase.GetPara(153, glngSys, , 0)
    gtype_UserSysParms.P163_项目执行前必须先收费或先记帐审核 = zldatabase.GetPara(163, glngSys, , 0)
    gtype_UserSysParms.P174_药品移库明确批次 = zldatabase.GetPara(174, glngSys, , 0)
    gtype_UserSysParms.P175_药品领用明确批次 = zldatabase.GetPara(175, glngSys, , 0)
    gtype_UserSysParms.P214_首次医嘱执行需要审核 = zldatabase.GetPara(214, glngSys, , 0)
    gtype_UserSysParms.P221_药品结存时点 = zldatabase.GetPara(221, glngSys, , 0)
    gtype_UserSysParms.P222_药房自动化发药接口 = zldatabase.GetPara(222, glngSys, , 0)
    gtype_UserSysParms.P240_药房处方审查 = zldatabase.GetPara(241, glngSys, , 0)
    gtype_UserSysParms.P241_处方审查时机 = zldatabase.GetPara(242, glngSys, , 0)
    gtype_UserSysParms.Para_输入方式 = zldatabase.GetPara(44, glngSys, , 11)
    gtype_UserSysParms.P275_零差价管理模式 = Val(zldatabase.GetPara(275, glngSys, , 0))
    gtype_UserSysParms.P213_中药配方每行中药味数 = Val(zldatabase.GetPara(213, glngSys, , 0))
    
    '取药品最大允许精度
    gstrSQL = "Select 零售金额, 成本价, 零售价, 实际数量 From 药品收发记录 Where Rownum < 1"
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "取药品精度")
    gtype_UserDrugDigits.Digit_金额 = rs.Fields(0).NumericScale
    gtype_UserDrugDigits.Digit_成本价 = rs.Fields(1).NumericScale
    gtype_UserDrugDigits.Digit_零售价 = rs.Fields(2).NumericScale
    gtype_UserDrugDigits.Digit_数量 = rs.Fields(3).NumericScale
    
    '取药品售价单位小数位数
    gstrSQL = "Select 内容, Nvl(精度, 0) 精度 From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 单位 = 1 "
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "取药品售价单位小数位数")
    
    If rs.RecordCount > 0 Then
        rs.Filter = "内容=1"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_成本价 = rs!精度
        
        rs.Filter = "内容=2"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_零售价 = rs!精度
        
        rs.Filter = "内容=3"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_数量 = rs!精度
        
        If gtype_UserSaleDigits.Digit_成本价 < 2 Or gtype_UserSaleDigits.Digit_成本价 > gtype_UserDrugDigits.Digit_成本价 Then
            gtype_UserSaleDigits.Digit_成本价 = gtype_UserDrugDigits.Digit_成本价
        End If
        
        If gtype_UserSaleDigits.Digit_零售价 < 2 Or gtype_UserSaleDigits.Digit_零售价 > gtype_UserDrugDigits.Digit_零售价 Then
            gtype_UserSaleDigits.Digit_零售价 = gtype_UserDrugDigits.Digit_零售价
        End If
        
        If gtype_UserSaleDigits.Digit_数量 < 2 Or gtype_UserSaleDigits.Digit_数量 > gtype_UserDrugDigits.Digit_数量 Then
            gtype_UserSaleDigits.Digit_数量 = gtype_UserDrugDigits.Digit_数量
        End If
    End If
    
    '公共全局参数
    gstrLike = IIf(Val(zldatabase.GetPara("输入匹配")) = 0, "%", "")
    gblnMyStyle = zldatabase.GetPara("使用个性化风格") = "1"
        
    '药品名称显示方式
    gint药品名称显示 = Val(zldatabase.GetPara("药品名称显示", , , 2))
    gint输入药品显示 = Val(zldatabase.GetPara("输入药品显示"))
    
    If gint药品名称显示 < 0 Or gint药品名称显示 > 2 Then gint药品名称显示 = 2
    If gint输入药品显示 < 0 Or gint输入药品显示 > 1 Then gint输入药品显示 = 0
    
    '简码方式
    gint简码方式 = Val(zldatabase.GetPara("简码方式"))
    If gint简码方式 < 0 Or gint简码方式 > 1 Then gint简码方式 = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function EsignIsOpen(ByVal lng部门ID As Long) As Boolean
    Dim rstemp As Recordset
    
    On Error GoTo errH
    gstrSQL = "select Zl_Fun_Getsignpar(5,[1]) 是否启用 from dual"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "电子签名使用部门", lng部门ID)

    If Not rstemp.EOF Then
        EsignIsOpen = (rstemp!是否启用 = 1)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer, Optional ByVal lng科室ID As Long) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strsql As String, intType As Integer
    Dim curDate As Date
    Dim intYear As Integer
    Dim PreFixNO As String  '年份前缀
    Dim strPre As String    '最大号码表中前2位
    Dim str编号 As String
    Dim dateCurDate As Date
    Dim intMonth As Integer
    Dim strMonth As String
    
    On Error GoTo errH
    
    dateCurDate = Sys.Currentdate
    intYear = Format(dateCurDate, "YYYY") - 1990
    PreFixNO = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
    intMonth = Month(dateCurDate)
    strMonth = intMonth
    strMonth = String(2 - Len(strMonth), "0") & strMonth
    
    If Len(strNo) >= 8 Then
        GetFullNO = Right(strNo, 8)
        Exit Function
    ElseIf Len(strNo) = 7 Then
        GetFullNO = PreFixNO & strNo
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNo, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNo
    
    strsql = "Select 编号规则,最大号码,Sysdate as 日期 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, "GetFullNO", intNum)
        
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
        strPre = Left(Nvl(rsTmp!最大号码, PreFixNO & "0"), 2)
    End If
    
    If intType = 0 Then
        '按年编号
        GetFullNO = strPre & Format(Right(strNo, 6), "000000")
    ElseIf intType = 1 Then
        '按日编号
        strsql = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strsql & Format(Right(strNo, 4), "0000")
    ElseIf intType = 2 Then
        '按科室分月编码
        gstrSQL = "Select 编号 From 科室号码表 Where 项目序号=[1] And Nvl(科室ID,0)=[2]"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "GetFullNO", intNum, lng科室ID)
        
        If rsTmp.RecordCount = 0 Then
            MsgBox "还未设置科室编号，无法产生号码！", vbInformation, gstrSysName
            Exit Function
        End If
        If Nvl(rsTmp!编号) = "" Then
            MsgBox "还未设置科室编号，无法产生号码！", vbInformation, gstrSysName
            Exit Function
        End If
        str编号 = Nvl(rsTmp!编号)
        
        '小于四位，按本月产生号码
        '五位或六位，则认为是指定月份的号码
        '七位，则认为是产生本年指定科室、月份的号码
        '大于等于八位，不处理
        If Len(strNo) <= 4 Then
            GetFullNO = PreFixNO & str编号 & strMonth & String(4 - Len(strNo), "0") & strNo
        ElseIf Len(strNo) <= 6 Then
            GetFullNO = String(6 - Len(strNo), "0") & GetFullNO
            GetFullNO = PreFixNO & str编号 & GetFullNO
        ElseIf Len(strNo) = 7 Then
            GetFullNO = PreFixNO & GetFullNO
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function


Public Function CopyNewRec(ByVal SourceRec As ADODB.Recordset) As ADODB.Recordset
        Dim RecTarget As New ADODB.Recordset
        Dim IntFields As Integer, LngLocate As Long
        '编制人:朱玉宝
        '编制日期:2000-11-02
        '该记录集与凭证控件对应
        '也使用于保存
        
        LngLocate = -1
        Set RecTarget = New ADODB.Recordset
        With RecTarget
                If .State = 1 Then .Close
                If SourceRec.RecordCount <> 0 Then
                        On Error Resume Next
                        err = 0
                        LngLocate = SourceRec.AbsolutePosition
                        If err <> 0 Then LngLocate = -1
                        SourceRec.MoveFirst
                End If
                For IntFields = 0 To SourceRec.Fields.count - 1
                        .Fields.Append SourceRec.Fields(IntFields).Name, SourceRec.Fields(IntFields).Type, SourceRec.Fields(IntFields).DefinedSize, adFldIsNullable     '0:表示新增
                Next
                
                .CursorLocation = adUseClient
                .CursorType = adOpenStatic
                .LockType = adLockOptimistic
                .Open
                
                If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
                Do While Not SourceRec.EOF
                        .AddNew
                        For IntFields = 0 To SourceRec.Fields.count - 1
                                .Fields(IntFields) = SourceRec.Fields(IntFields).Value
                        Next
                        .Update
                        SourceRec.MoveNext
                Loop
        End With
        
        If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
        If LngLocate > 0 Then SourceRec.Move LngLocate - 1
        Set CopyNewRec = RecTarget
End Function




Public Function GetUserInfo() As Boolean
    Dim rsUser As ADODB.Recordset
    
    Set rsUser = Sys.GetUserInfo
    
    With rsUser
        If Not .EOF Then
            glngUserId = !Id '当前用户id
            UserInfo.用户ID = !Id
            gstrUserCode = !编号 '当前用户编码
            UserInfo.用户编码 = !编号
            gstrUserName = IIf(IsNull(!姓名), "", !姓名) '当前用户姓名
            UserInfo.用户姓名 = IIf(IsNull(!姓名), "", !姓名)
            gstrUserAbbr = IIf(IsNull(!简码), "", !简码) '当前用户简码
            UserInfo.用户简码 = IIf(IsNull(!简码), "", !简码)
            glngDeptId = !部门ID '当前用户部门id
            UserInfo.部门ID = !部门ID
            gstrDeptCode = !部门码 '当前用户
            UserInfo.部门编码 = !部门码
            gstrDeptName = !部门名 '当前用户
            UserInfo.部门名称 = !部门名
            GetUserInfo = True
        Else
            glngUserId = 0 '当前用户id
            gstrUserCode = "" '当前用户编码
            gstrUserName = "" '当前用户姓名
            gstrUserAbbr = "" '当前用户简码
            glngDeptId = 0 '当前用户部门id
            gstrDeptCode = "" '当前用户
            gstrDeptName = "" '当前用户
            
            
            UserInfo.用户ID = 0
            UserInfo.用户编码 = ""
            UserInfo.用户姓名 = ""
            UserInfo.用户简码 = ""
            UserInfo.部门ID = 0
            UserInfo.部门编码 = ""
            UserInfo.部门名称 = ""
        End If
    End With
End Function
Public Function GetUnit(ByVal lng药房id As Long, ByVal Int单据 As Integer, ByVal strNo As String, ByVal int门诊标志 As Integer) As String
    '返回指定库房、单据、NO适用的药品单位
    Dim intUnit As Integer
    Dim blnMoved As Boolean
    Dim rstemp As New ADODB.Recordset
    
    '根据系统参数设定的单位显示数据
    intUnit = Val(zldatabase.GetPara("药房属性", glngSys, 1341, 0))
    If intUnit = 0 Then
        intUnit = int门诊标志
    End If
    If intUnit = 1 Or intUnit = 4 Then
        GetUnit = GetSpecUnit(lng药房id, gint门诊药房)
    Else
        GetUnit = GetSpecUnit(lng药房id, gint住院药房)
    End If
End Function
Public Function GetSpecUnit(ByVal lng库房id As Long, ByVal int范围 As Integer) As String
    Dim strobjTemp As String                    '保存服务对象字符串
    Dim strWorkTemp As String                   '保存工作性质字符串
    Dim strUnit As String
    Dim rsProperty As New ADODB.Recordset
    Dim strsql As String
    
    '返回指定库房指定适用范围的单位
    On Error GoTo ErrHand
    
    gstrSQL = "Select Nvl(性质,1) AS 单位 From 药品库房单位 Where 库房ID=[1] And 适用范围=[2] "
    Set rsProperty = zldatabase.OpenSQLRecord(gstrSQL, "提取单位", lng库房id, int范围)
   
    If rsProperty.RecordCount = 1 Then
        strUnit = rsProperty!单位
    Else
        gstrSQL = "SELECT distinct 服务对象,工作性质 From 部门性质说明 Where 部门ID =[1]"
        Set rsProperty = zldatabase.OpenSQLRecord(gstrSQL, "读取药品单位", lng库房id)
    
        '取服务对象及部门性质
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            '住院单位
            strUnit = 3
        ElseIf InStr(strobjTemp, "1") <> 0 Then
            '门诊单位
            strUnit = 2
        ElseIf InStr(strWorkTemp, "药库") <> 0 Then
            '药库单位
            strUnit = 4
        Else
            '售价单位：主要是制剂室
            strUnit = 1
        End If
    End If
    
    '转换为真实的单位返回给调用者
    GetSpecUnit = Switch(strUnit = 1, "售价单位", strUnit = 2, "门诊单位", strUnit = 3, "住院单位", strUnit = 4, "药库单位")
    If glngSys / 100 = 8 Then
        '药店只有售价单位与药库单位
        GetSpecUnit = IIf(strUnit = 1, "售价单位", "药库单位")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

'取药品单位名称
Public Function GetDrugUnit(ByVal lng库房id As Long, ByVal frmCaption As String, Optional ByVal bln处方 As Boolean = True) As String
    Dim rsProperty As New Recordset
    Dim strobjTemp As String                    '保存服务对象字符串
    Dim strWorkTemp As String                   '保存工作性质字符串
    Dim intUnit As Integer, strUnit As String
    Dim bln缺省 As Boolean
    Dim lngModul As Long
    
    On Error GoTo ErrHand
    
    If frmCaption Like "药品申领管理*" Then
        lngModul = 1343
    ElseIf frmCaption Like "协定药品入库*" Then
        lngModul = 1344
    ElseIf frmCaption Like "药品移库管理*" Then
        lngModul = 1304
    End If
    
    intUnit = 0
    '如果是申领单，则直接返回注册表中的单位
    If lngModul = 1343 Or lngModul = 1304 Or lngModul = 1344 Then
        intUnit = Val(zldatabase.GetPara("药品单位", glngSys, lngModul))
        '本地参数设置的单位顺序如下：0-缺省;1-药库;2-门诊;3-住院;4-售价，需要转换为与系统参数的一致
        If intUnit = 1 Then
            intUnit = 4
        ElseIf intUnit = 4 Then
            intUnit = 1
        End If
        strUnit = intUnit
    End If
    
    If intUnit = 0 Then
        gstrSQL = "SELECT distinct 服务对象,工作性质 From 部门性质说明 Where 部门ID =[1]"
        Set rsProperty = zldatabase.OpenSQLRecord(gstrSQL, "读取药品单位", lng库房id)
        
        '取服务对象及部门性质
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        
        If InStr(strWorkTemp, "药库") <> 0 Then
            '药库单位
            intUnit = 1
            strUnit = 4
        ElseIf InStr(strobjTemp, "1") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            '门诊单位
            intUnit = 2
            strUnit = 2
        ElseIf InStr(strobjTemp, "2") <> 0 Then
            '住院单位
            intUnit = 3
            strUnit = 3
        Else
            '售价单位：主要是制剂室
            intUnit = 4
            strUnit = 1
        End If
        
        '取该药房缺省该使用的单位
        GetDrugUnit = GetSpecUnit(lng库房id, intUnit)
    Else
        GetDrugUnit = Switch(strUnit = 1, "售价单位", strUnit = 2, "门诊单位", strUnit = 3, "住院单位", strUnit = 4, "药库单位")
    End If
    
    '转换为真实的单位返回给调用者
    
    If glngSys / 100 = 8 Then
        '药店只有售价单位与药库单位
        GetDrugUnit = IIf(strUnit = 1, "售价单位", "药库单位")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDrugUnit = "售价单位"
End Function



'按编码，名称，别名查找某一列
Public Function FindRow(ByVal mshBill As BillEdit, ByVal int比较列 As Integer, _
    ByVal str比较值 As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim StrCode As String
    Dim rsCode As New Recordset
    
    On Error GoTo errHandle
    FindRow = True
    With mshBill
        If .rows = 2 Then Exit Function
        If str比较值 = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                StrCode = .TextMatrix(intRow, int比较列)
                If InStr(1, UCase(StrCode), UCase(str比较值)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int比较列
                    .SetRowColor CLng(intRow), &HFFCECE, True
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = " SELECT DISTINCT b.编码 " & _
                  " FROM " & _
                  "    (SELECT DISTINCT A.收费细目id " & _
                  "    FROM 收费项目别名 A" & _
                  "    Where A.简码 LIKE [1]) a," & _
                  " 收费项目目录 B " & _
                  " Where a.收费细目id = b.ID"
        Set rsCode = zldatabase.OpenSQLRecord(gstrSQL, "查找药品", IIf(gstrMatchMethod = "0", "%", "") & str比较值 & "%")
        
        If rsCode.EOF Then
            FindRow = False
            Exit Function
        End If
        
        For intRow = intStartRow To .rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                StrCode = .TextMatrix(intRow, int比较列)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(StrCode), UCase(rsCode!编码)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int比较列
                        .SetRowColor CLng(intRow), &HFFCECE, True
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            
            End If
        Next
        rsCode.Close
    End With
    FindRow = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetLength(ByVal strTable As String, ByVal strColumn As String) As Integer
    Dim rsPar As New ADODB.Recordset
    '获取指定表特定字段的长度
    
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    GetLength = 0
    
    With rsPar
        gstrSQL = "Select " & strColumn & " From " & strTable & " Where Rownum<1"
        Call zldatabase.OpenRecordset(rsPar, gstrSQL, "获取长度")
        
        If err <> 0 Then
            MsgBox "数据表[" & strTable & "]不存在，请与开发商联系！", vbInformation, gstrSysName
        End If
        GetLength = .Fields(0).DefinedSize
        .Close
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ReturnSQL(ByVal lng库房id As Long, ByVal strCaption As String, _
    Optional ByVal bln调拨 As Boolean = True, _
    Optional ByVal lngModuleNO As Long = 0) As ADODB.Recordset
    
    Dim str库房性质 As String, str药品流向 As String, str站点限制 As String, strsql As String
    '根据药品流向控制表的数据，提取对方库房
    'Writed by zyb
    '-----------------调拨-----------------
    '所在库房是当前库房的，提取流向 In (1"可流向对方库房",3"可双向流通")
    '对方库房是当前库房的，提取流向 IN (2"可流向所在库房",3"可双向流通")
    '-----------------申领-----------------
    '所在库房是当前库房的，提取流向 In (2"可流向所在库房",3"可双向流通")
    '对方库房是当前库房的，提取流向 IN (1"可流向对方库房",3"可双向流通")
    
    On Error GoTo errHandle
    str站点限制 = GetDeptStationNode(lng库房id)
    str库房性质 = "('H','I','J','K','L','M','N')"
    
    str药品流向 = ",(Select 对方库房ID ID From 药品流向控制" & _
            " Where 所在库房ID=[1] And 流向 In (" & IIf(bln调拨, 1, 2) & ",3)" & _
            " Union" & _
            " Select 所在库房ID ID From 药品流向控制" & _
            " Where 对方库房ID=[1] And 流向 In (" & IIf(bln调拨, 2, 1) & ",3)) D"
    Select Case lngModuleNO
        Case 1343   '药品申领管理
            strsql = " SELECT DISTINCT a.id,a.编码,a.名称, Decode(Instr(',H,I,J,', ',' || b.编码 || ','), 0, 0, 1) As 药库性质 " & _
                    " FROM 部门性质说明 c, 部门性质分类 b, 部门表 a" & str药品流向 & _
                    " Where c.工作性质 = b.名称" & _
                    " AND b.编码||'' in " & str库房性质 & _
                    " AND a.id = c.部门id And A.ID=D.ID " & _
                    " AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
                    " Order by a.编码"
        Case Else
            strsql = " SELECT DISTINCT a.id,a.编码,a.名称, Decode(Instr(',H,I,J,', ',' || b.编码 || ','), 0, 0, 1) As 药库性质 " & _
                    " FROM 部门性质说明 c, 部门性质分类 b, 部门表 a" & str药品流向 & _
                    " Where c.工作性质 = b.名称" & _
                    " AND b.编码||'' in " & str库房性质 & _
                    " AND a.id = c.部门id And A.ID=D.ID" & IIf(str站点限制 <> "", " AND a.站点=[2] ", "") & _
                    " AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
                    " Order by a.编码"
    End Select
    Set ReturnSQL = zldatabase.OpenSQLRecord(strsql, strCaption, lng库房id, str站点限制)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckRepeatMedicine(ByVal MyBill As Object, ByVal strDrugInfo As String, ByVal intExceptRow As Integer) As Boolean
    '药品流通编辑界面检查录入的药品是否重复
    'MyBill：表单控件（药品列表）
    'strDrugInfo：药品ID，批次及对应列号（格式：药品ID,药品ID列号|批次,批次列号）
    'intExceptRow：排除指定的行（不检查这一行）
    Dim n As Integer
    Dim lng药品id As Long
    Dim int药品ID列号 As Integer
    Dim lng批次 As Long
    Dim int批次列号 As Integer
    
    On Error GoTo errHandle
    lng药品id = Val(Split(Split(strDrugInfo, "|")(0), ",")(0))
    int药品ID列号 = Val(Split(Split(strDrugInfo, "|")(0), ",")(1))
    lng批次 = Val(Split(Split(strDrugInfo, "|")(1), ",")(0))
    int批次列号 = Val(Split(Split(strDrugInfo, "|")(1), ",")(1))
    
    With MyBill
        For n = 1 To .rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If n <> intExceptRow And Val(.TextMatrix(n, int药品ID列号)) = lng药品id And Val(.TextMatrix(n, int批次列号)) = lng批次 Then
                    MsgBox "对不起，已有该药品或该药品的相同批次，不能重复输入！", vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    CheckRepeatMedicine = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function












Public Sub CheckLapse(ByVal str效期 As String)
    '失效药品检查
    If Not IsDate(str效期) Then Exit Sub
    If Format(str效期, "yyyy-MM-dd") < Format(Sys.Currentdate, "yyyy-MM-dd") Then
        MsgBox "该药品已经失效了！", vbInformation, gstrSysName
    End If
End Sub

Public Sub zlPlugIn_Ini(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object)
    '外挂扩展接口初始化
    If objPlugIn Is Nothing Then
        On Error Resume Next
        Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not objPlugIn Is Nothing Then
            Call objPlugIn.Initialize(gcnOracle, lngSys, lngModul)
            If InStr(",438,0,", "," & err.Number & ",") = 0 Then
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub


Public Sub zlPlugIn_SetMenu(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object, _
    cbcMain As CommandBarControls, ByVal lngMenuPlugInMain As Long)
    '设置扩展功能的菜单项目
    '参数：lngSys-系统，lngModul-模块号，objPlugIn-扩展外挂对象，cbcMain-CommandBar主菜单对象，lngMenuPlugInMain-外挂菜单
    Dim strFunc As String, strFuncName As String '记录扩展功能
    Dim lngFuncID As Long
    Dim blnGroup As Boolean
    Dim intToolbarCount As Integer
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim i As Integer
    
    If Not objPlugIn Is Nothing Then
        On Error Resume Next
        
        '外挂部件有扩展功能
        strFunc = objPlugIn.GetFuncNames(lngSys, lngModul)
        If InStr(",438,0,", "," & err.Number & ",") = 0 Then
            MsgBox "zlPlugIn 外挂部件执行 GetFuncNames 时出错：" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
        End If
        
        err.Clear: On Error GoTo 0
    End If
    
    If strFunc = "" Then Exit Sub
    
    With cbcMain
        Set objPopup = .Add(xtpControlButtonPopup, lngMenuPlugInMain, "扩展(&E)")
        objPopup.BeginGroup = True
        
        With objPopup.CommandBar.Controls
            For i = 0 To UBound(Split(strFunc, ","))
                lngFuncID = lngMenuPlugInMain + i + 1
                strFuncName = Split(strFunc, ",")(i)
                
                blnGroup = InStr(strFuncName, "|") > 0
                strFuncName = Replace(strFuncName, "InTool:", "")
                strFuncName = Replace(strFuncName, "|:", "")
                
                Set cbrControl = .Add(xtpControlButton, lngFuncID, strFuncName)
                If i <= 9 Then cbrControl.Caption = cbrControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                cbrControl.IconId = lngMenuPlugInMain
                cbrControl.Parameter = strFuncName
                cbrControl.BeginGroup = blnGroup
            Next
        End With
    End With
End Sub

Public Sub zlPlugIn_SetToolbar(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object, _
    cbrToolBar As CommandBarControls, ByVal lngMenuPlugInMain As Long)
    '设置扩展功能的工具栏项目
    '参数：lngSys-系统，lngModul-模块号，objPlugIn-扩展外挂对象，cbrToolBar-CommandBar工具栏对象，lngMenuPlugInMain-外挂菜单
    Dim strFunc As String, strFuncName As String '记录扩展功能
    Dim lngFuncID As Long
    Dim blnGroup As Boolean
    Dim intToolbarCount As Integer
    Dim cbrControl As CommandBarControl
    Dim i As Integer
    
    If Not objPlugIn Is Nothing Then
        On Error Resume Next
        
        '外挂部件有扩展功能
        strFunc = objPlugIn.GetFuncNames(lngSys, lngModul)
        If InStr(",438,0,", "," & err.Number & ",") = 0 Then
            MsgBox "zlPlugIn 外挂部件执行 GetFuncNames 时出错：" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
        End If
        
        err.Clear: On Error GoTo 0
    End If
    
    If strFunc = "" Then Exit Sub
    
    With cbrToolBar
        For i = 0 To UBound(Split(strFunc, ","))
            strFuncName = Split(strFunc, ",")(i)
            lngFuncID = lngMenuPlugInMain + i + 1
            
            If InStr(strFuncName, "InTool:") > 0 Then
                intToolbarCount = intToolbarCount + 1
                blnGroup = (intToolbarCount = 1 Or InStr(strFuncName, "|") > 0)
                strFuncName = Replace(strFuncName, "InTool:", "")
                strFuncName = Replace(strFuncName, "|:", "")
    
                Set cbrControl = .Add(xtpControlButton, lngFuncID, strFuncName)
                cbrControl.IconId = lngMenuPlugInMain
                cbrControl.Parameter = strFuncName
                cbrControl.BeginGroup = blnGroup
            End If
        Next
    End With
End Sub


Public Sub zlPlugIn_Unload(objPlugIn As Object)
    '卸载外挂接口
    Set objPlugIn = Nothing
End Sub


Public Function 检查库存数据(ByVal lng库房id As Long, ByVal lng药品id As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    Dim bln库存是否分批 As Boolean, bln分批 As Boolean, bln库房 As Boolean
    '通过药品选择器输入药品时，如果药品库存中的数据与从部门性质、药品目录中的分批属性判断出的不一致，则报错
    On Error GoTo errHandle
    检查库存数据 = False
    
    '如果没有库存记录，则直接退出
    gstrSQL = " Select Count(*) 记录数 From 药品库存 " & _
              " Where 库房ID=[1] And 性质=1 And 药品ID=[2]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "是否存在库存数据", lng库房id, lng药品id)
    
    If rsCheck!记录数 = 0 Then
        检查库存数据 = True
        Exit Function
    End If
    
    '存在分批记录则表明分批
    gstrSQL = " Select Count(*) 分批 From 药品库存 " & _
              " Where 库房ID=[1] And 性质=1 And Nvl(批次,0)<>0 And 药品ID=[2]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "检查库存数据", lng库房id, lng药品id)
    
    bln库存是否分批 = (rsCheck!分批 <> 0)
    
    '先判断是否是库房
    gstrSQL = "select 部门ID from 部门性质说明 where (工作性质 like '%药房' Or 工作性质 like '%制剂室') And 部门id=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "取部门性质", lng库房id)
    
    bln库房 = (rsCheck.EOF)
        
    '判断对应的药品目录中的分批属性
    gstrSQL = " Select Nvl(药库分批,0) 分批核算,nvl(药房分批,0) 药房分批核算 " & _
              " From 药品规格 Where 药品ID=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "取药品目录中的分批属性", lng药品id)
              
    If bln库房 Then
        bln分批 = (rsCheck!分批核算 = 1)
    Else
        bln分批 = (rsCheck!药房分批核算 = 1)
    End If
    
    检查库存数据 = (bln库存是否分批 = bln分批)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




'取药品金额、价格和数量的小数位数
Public Function GetDigit(ByVal int性质 As Integer, ByVal int类别 As Integer, ByVal int内容 As Integer, Optional ByVal int单位 As Integer) As Integer
    'int性质：0-计算精度;1-显示精度
    'int类别：1-药品;2-卫材
    'int内容：1-成本价;2-零售价;3-数量;4-金额
    'int单位：如果是取金额位数，可以不输入该参数
    '         药品单位:1-售价;2-门诊;3-住院;4-药库;
    '         卫材单位:1-散装;2-包装
    '返回：最小2，最大为数据库最大小数位数
    
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If int内容 = 4 Then
        int单位 = 5
    End If
    gstrSQL = "Select Nvl(精度, 0) 精度 From 药品卫材精度 Where 性质 = [1] And 类别 = [2] And 内容 = [3] And 单位 = [4] "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "取药品" & Choose(int内容, "成本价", "零售价", "数量") & "小数位数", int性质, int类别, int内容, int单位)
    
    If rsTmp.RecordCount > 0 Then
        GetDigit = rsTmp!精度
    End If
    
    If GetDigit = 0 Then
        '如果没有设置精度，则取数据库允许的最大位数
        GetDigit = Choose(int内容, gtype_UserDrugDigits.Digit_成本价, gtype_UserDrugDigits.Digit_零售价, gtype_UserDrugDigits.Digit_数量, gtype_UserDrugDigits.Digit_金额)
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDigit = Choose(int内容, gtype_UserDrugDigits.Digit_成本价, gtype_UserDrugDigits.Digit_零售价, gtype_UserDrugDigits.Digit_数量, gtype_UserDrugDigits.Digit_金额)
End Function

'根据库房的包装单位来取药品的价格、数量、金额小数位数
Public Sub GetDrugDigit(ByRef lng库房id As Long, ByVal frmCaption As String, ByRef intUnit As Integer, ByRef intCostDigit As Integer, ByRef intPricedigit As Integer, ByRef intNumberDigit As Integer, ByRef intMoneyDigit As Integer)
    Dim strUnit As String
    
    Const conInt计算精度 As Integer = 0
    
    Const conInt药品 As Integer = 1
    
    Const conint售价单位 As Integer = 1
    Const conint门诊单位 As Integer = 2
    Const conint住院单位 As Integer = 3
    Const conint药库单位 As Integer = 4
        
    Const conInt成本价 As Integer = 1
    Const conInt售价 As Integer = 2
    Const conInt数量 As Integer = 3
    Const conInt金额 As Integer = 4
    
    strUnit = GetDrugUnit(lng库房id, frmCaption)
    
    Select Case strUnit
        Case "售价单位"             '售价单位：主要是制剂室
            intUnit = conint售价单位
        Case "门诊单位"
            intUnit = conint门诊单位
        Case "住院单位"
            intUnit = conint住院单位
        Case "药库单位"
            intUnit = conint药库单位
    End Select

    '分别取药品成本价、售价、数量、金额的小数位数
    intCostDigit = GetDigit(conInt计算精度, conInt药品, conInt成本价, intUnit)
    intPricedigit = GetDigit(conInt计算精度, conInt药品, conInt售价, intUnit)
    intNumberDigit = GetDigit(conInt计算精度, conInt药品, conInt数量, intUnit)
    intMoneyDigit = GetDigit(conInt计算精度, conInt药品, conInt金额)

End Sub



Public Function 药品单据审核(ByVal str填制人 As String) As Boolean
    '药品单据审核时，是否判断审核人与填制人，其返回审核结果
    Dim blnBillVerify As Boolean
    Dim rsSystemPara As New Recordset
    Dim intTemp As Integer
    
    On Error GoTo errHandle
    
    药品单据审核 = True
    
    intTemp = Val(zldatabase.GetPara(64, glngSys, 0))
    If intTemp = 0 Then
        blnBillVerify = False
        Exit Function
    Else
        blnBillVerify = True
    End If
    
    If Not blnBillVerify Then Exit Function
    
    药品单据审核 = (Trim(str填制人) <> Trim(gstrUserName))
    If Not 药品单据审核 Then MsgBox "填制人与审核人不能是同一人，请检查！", vbInformation, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBillInfo(ByVal lng单据 As Long, ByVal strNo As String, Optional ByVal bln填制日期 As Boolean = True) As String
    Dim rsBillInfo As New ADODB.Recordset
    '获取单据的最大修改时间
    
    On Error GoTo errHandle
    gstrSQL = " Select to_char(Max(" & IIf(bln填制日期, "填制日期", "审核日期") & "),'yyyyMMddhh24miss') 日期 From 药品收发记录 " & _
            " Where 单据=[1] And NO=[2]"
    Set rsBillInfo = zldatabase.OpenSQLRecord(gstrSQL, "获取单据的最大修改时间", lng单据, strNo)
    
    With rsBillInfo
        '返回空，表示已经删除
        If .EOF Then Exit Function
        If IsNull(!日期) Then Exit Function
        GetBillInfo = !日期
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function 检查单价(ByVal lng单据 As Long, ByVal strNo As String, Optional ByVal blnmsg As Boolean = True, Optional ByVal bln移库单 As Boolean = False) As Boolean
    Dim rsPrice As New ADODB.Recordset
    Dim lng药品_Last As Long, lng药品_Cur As Long
    Dim intPricedigit As Integer
    Dim intCostDigit As Integer
    '检查药品的价格是否为最新的价格（按药库单位进行比较），允许继续操作
    '由于在保存前判断很麻烦，且各种单据的表格中保存的数据不一样，因此，待保存完成之后且提交前对已保存的数据进行检查
    '药品相同的记录略过
    
    '自动批量检查并执行调价
    On Error GoTo errHandle
    
    Call AutoAdjustPrice_ByNO(lng单据, strNo)
    intPricedigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
    
    gstrSQL = " Select '售价' As 类型, a.序号, a.药品id , 0 原价, b.现价" & _
            " From 药品收发记录 A," & _
                 " (Select 收费细目id, Nvl(现价, 0) 现价, 执行日期" & _
                   " From 收费价目" & _
                   " Where (终止日期 Is Null Or Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, 收费项目目录 C" & _
            " Where a.单据 = [1] And a.No = [2] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & intPricedigit & ") <> Round(b.现价, " & intPricedigit & ") And" & _
              "    NVL(c.是否变价, 0) = 0 " & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id , 0 原价, decode(x.现价,null,decode(nvl(b.零售价,0),0,b.实际金额 / b.实际数量,b.零售价),x.现价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C ," & _
            "      (Select x.药品id,x.库房id,x.批次,x.现价 from 药品价格记录 x where x.价格类型 = 1 and (x.终止日期 Is Null Or Sysdate Between x.执行日期 And Nvl(x.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.单据 = [1] And a.No = [2] And c.Id = a.药品id And Round(a.零售价," & intPricedigit & ") <> Round(decode(x.现价,null,decode(nvl(b.零售价,0),0,b.实际金额 / b.实际数量,b.零售价),x.现价), " & intPricedigit & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
                  " AND a.药品id = x.药品id(+) And a.库房id = x.库房id(+) And Nvl(a.批次, 0) = Nvl(x.批次(+), 0) " & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id , 0 原价, decode(x.现价,null,b.平均成本价,x.现价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B ," & _
            "      (Select x.药品id,x.库房id,x.批次,x.现价 from 药品价格记录 x where x.价格类型 = 2 and (x.终止日期 Is Null Or Sysdate Between x.执行日期 And Nvl(x.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.单据 = [1] And a.No = [2] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & intCostDigit & ")<>round(decode(x.现价,null,b.平均成本价,x.现价)," & intCostDigit & ") And a.库房id = b.库房id and a.入出系数=-1  and b.性质=1" & _
            " AND a.药品id = x.药品id(+) And a.库房id = x.库房id(+) And Nvl(a.批次, 0) = Nvl(x.批次(+), 0) " & _
            " Order By 类型, 药品id, 序号"
    Set rsPrice = zldatabase.OpenSQLRecord(gstrSQL, "取当前价格", lng单据, strNo)
            
    If rsPrice.EOF Then
        检查单价 = True
        Exit Function
    End If
    
    lng药品_Last = 0
    With rsPrice
        Do While Not .EOF
            lng药品_Cur = !药品ID
            If lng药品_Cur <> lng药品_Last Then
                If blnmsg Then
                    MsgBox "第" & IIf(bln移库单, Round(!序号 / 2 + 0.49), !序号) & "行药品的" & !类型 & "不是最新价格，将按照最新价格更新界面！", vbInformation, gstrSysName
                    Exit Function
                Else
                    Exit Function
                End If
            End If
            
            lng药品_Last = lng药品_Cur
            .MoveNext
        Loop
        检查单价 = True
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DepotProperty(ByVal lng人员id As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '返回指定人员是否具有药库性质
    On Error GoTo errHandle
    gstrSQL = "Select Distinct 工作性质 From 部门人员 B,部门性质说明 A " & _
             " Where A.工作性质 like '%药库' And " & _
             " A.部门id = B.部门id And B.人员id = [1]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "取部门性质", lng人员id)
    If rsCheck.RecordCount <> 0 Then
        DepotProperty = True
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowCostPrice() As Boolean
    '药库人员不管，只管药房人员，以参数控制为准
    If DepotProperty(glngUserId) Then
        ShowCostPrice = True
    Else
        ShowCostPrice = (gtype_UserSysParms.P85_药房查看单据成本价 = 1)
    End If
End Function
Public Function IsOwner(ByVal strUser As String) As Boolean
    Dim rstemp As New ADODB.Recordset
    '判断传入的用户是不是所有者或DBA用户
    On Error GoTo errHandle
    gstrSQL = "SELECT 1 FROM DUAL " & _
            " WHERE EXISTS(SELECT 1 FROM ZLSYSTEMS WHERE 所有者=[1])"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "判断该用户是不是所有者", UCase(strUser))
    IsOwner = (rstemp.RecordCount <> 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Is中药库房(ByVal lng药房id As Long) As Boolean
    Dim rstemp As New ADODB.Recordset
    Dim str库房性质 As String
    
    gstrSQL = "Select 工作性质 From 部门性质说明 Where 部门id =[1]"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "判断库房性质", lng药房id)
    Do While Not rstemp.EOF
        str库房性质 = str库房性质 & "," & rstemp!工作性质
        rstemp.MoveNext
    Loop
    If str库房性质 Like "*中药*" Then
        Is中药库房 = True
    Else
        Is中药库房 = False
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsLowerLimit(ByVal lng库房id As Long, ByVal lng药品id As Long) As Boolean
    '判断该药品在当前库存的库存是否低于库存下限，是则返回真
    Dim dbl库存数量 As Double, dbl下限 As Double
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '提取库存数量
    gstrSQL = " Select Sum(Nvl(实际数量,0)) AS 库存数量 From 药品库存" & _
              " Where 性质=1 And 库房ID=[1] And 药品ID=[2]"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取指定库房的实际库存", lng库房id, lng药品id)
              
    If rstemp.RecordCount = 1 Then dbl库存数量 = Nvl(rstemp!库存数量, 0)
    
    '提取储备限额中的下限
    gstrSQL = " Select Nvl(下限,0) AS 下限 From 药品储备限额" & _
              " Where 库房ID=[1] And 药品ID=[2]"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取储备限额中的下限", lng库房id, lng药品id)
    
    If rstemp.RecordCount = 1 Then dbl下限 = rstemp!下限
    
    IsLowerLimit = (dbl库存数量 < dbl下限)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsReceiptBalance_Charge(ByVal intType As Integer, ByVal str权限 As String, ByVal lng单据 As Long, ByVal strNo As String, ByVal str序号 As String, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer, Optional ByVal lngModle As Long) As Boolean
    'intType    ：0-发药;1-退药
    'str权限    ：当前操作员拥有的权限
    'lng单据    ：当前单据类型
    'strNO      ：当前单据号
    'str序号    ：费用序号
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If lng单据 = 8 Then
        IsReceiptBalance_Charge = True
        Exit Function
    End If
    
    '发药、退药状态分别检查是否有权限“发已结帐处方”和“退已结帐处方”，检查该处方是否已结帐，已结帐处方不允许发退药操作
    If (intType = 0 And InStr(1, str权限, "发已结帐处方") = 0) Or (intType = 1 And InStr(1, str权限, "退已结帐处方") = 0) Then
        '合并门诊、住院费用记录，按结账金额倒序排序
        gstrSQL = "Select Nvl(Sum(Nvl(结帐金额,0)),0) AS 结帐金额   " & _
                 "  From 门诊费用记录   " & _
                 "  Where Instr([1], ',' || 序号 || ',') > 0 " & _
                 "  And Mod(记录性质,10) = 2 and NO = [2] "
        If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
        Else
            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        End If
        gstrSQL = gstrSQL & " Order By 结帐金额 Desc"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否已结帐", "," & str序号 & ",", strNo)
        
        If Nvl(rstemp!结帐金额, 0) <> 0 Then
            If lngModle = 1 Then
                MsgBox "病人已结帐，你没有对已结帐病人的输液单进行销帐审核的权限，操作中止！", vbInformation, gstrSysName
            ElseIf lngModle = 2 Then
                MsgBox "病人已结帐，你没有对已结帐病人的输液单进行摆药的权限，操作中止！", vbInformation, gstrSysName
            ElseIf lngModle = 3 Then
                MsgBox "病人已结帐，你没有对已结帐病人的输液单进行取消摆药的权限，操作中止！", vbInformation, gstrSysName
            ElseIf lngModle = 4 Then
                MsgBox "病人已结帐，你没有对已结帐病人的输液单进行取消配药的权限，操作中止！", vbInformation, gstrSysName
            Else
                MsgBox "在处方[" & strNo & "]病人已结账，你没有对已结账病人的处方进行发药、退药的权限，操作中止！", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    End If
    
    IsReceiptBalance_Charge = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsOutPatient(ByVal str权限 As String, ByVal lng单据 As Long, ByVal strNo As String, _
    ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer, Optional ByVal lng病人ID As Long, _
    Optional ByVal lng主页ID As Long, Optional ByVal lngModle As Long, Optional ByVal str姓名 As String) As Boolean
    '功能说明：如果当前病人是住院病人，如果没有权限“发退出院病人处方”，则不允许发退药操作
    Const str发退出院病人处方 As String = "发退出院病人处方"
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    '如果有权限“发退出院病人处方”，则允许发退药操作
    If InStr(1, str权限, str发退出院病人处方) > 0 Then
        IsOutPatient = True
        Exit Function
    End If
    
    '如果当前病人是住院病人，如果没有权限“发退出院病人处方”，则不允许发退药操作
    '如果未传入病人ID，则自动提取
    If lng病人ID = 0 Then
'        gstrSQL = "Select A.病人ID,c.主页id From 门诊费用记录 A, 药品收发记录 B,病人医嘱记录 C Where A.ID = B.费用ID  And A.医嘱序号=C.id And b.单据 = [1] And b.No = [2] And Rownum = 1 "
'
'        If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
'        Else
'            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
'        End If
        
        gstrSQL = "Select a.病人id, c.主页id" & vbNewLine & _
                "From 门诊费用记录 A, 药品收发记录 B, 病人医嘱记录 C" & vbNewLine & _
                "Where a.Id = b.费用id And a.医嘱序号 = c.Id And b.单据 = [1] And b.No = [2] And Rownum = 1" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select d.病人id, d.主页id" & vbNewLine & _
                "From 住院费用记录 D, 药品收发记录 B" & vbNewLine & _
                "Where d.Id = b.费用id And b.单据 = [1] And b.No = [2] And Rownum = 1"

        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人ID", lng单据, strNo)
        
        '特殊情况，找不到病人ID则不进行下一步检查
        If rstemp.EOF Then
            IsOutPatient = True
            Exit Function
        End If
        
        lng病人ID = rstemp!病人ID
        lng主页ID = Nvl(rstemp!主页id, 0)
    End If
    
    '取病人姓名
    If str姓名 = "" Then
        gstrSQL = "Select 姓名 From 病人信息 Where 病人ID=[1]"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人姓名", lng病人ID)

        str姓名 = rstemp!姓名
    End If
    
    '检查病人已预出院或出院
    gstrSQL = " Select 1 From 病案主页" & _
              " Where 病人ID=[1] and 主页id=[2] " & _
              " And (出院日期 Is Not NULL)"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否已出院", lng病人ID, lng主页ID)
    
    If rstemp.RecordCount <> 0 Then
        If lngModle = 1 Then
            MsgBox "病人“" & str姓名 & "”已出院，你没有对已出院病人的输液单进行销帐审核的权限，操作中止！", vbInformation, gstrSysName
        ElseIf lngModle = 2 Then
            MsgBox "病人“" & str姓名 & "”已出院，你没有对已出院病人的输液单进行摆药的权限，操作中止！", vbInformation, gstrSysName
        ElseIf lngModle = 3 Then
            MsgBox "病人“" & str姓名 & "”已出院，你没有对已出院病人的输液单进行取消摆药的权限，操作中止！", vbInformation, gstrSysName
        ElseIf lngModle = 4 Then
            MsgBox "病人“" & str姓名 & "”已出院，你没有对已出院病人的输液单进行取消配药的权限，操作中止！", vbInformation, gstrSysName
        Else
            MsgBox "在处方[" & strNo & "]中，病人“" & str姓名 & "”已出院，你没有对已出院病人的处方进行发药、退药的权限，操作中止！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    IsOutPatient = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Calc_Clique(ByVal lng药品id As Long, ByVal dbl申领数量 As Double, Optional ByVal bln校验 As Boolean = False) As Double
    Dim dbl实际数量 As Double
    Dim dbl商 As Double, dbl余 As Double, dbl阀值 As Double
    Dim rstemp As New ADODB.Recordset
    '根据申领阀值计算得出实际数量，当分批药品计算时，传入的申领数量可能就是库存数量，此时校验参数为真，计算出的结果不能大于申领数量，也就是库存数量
    '如果传入的正确的，则肯定不需校正（应用于申领）
'    On Error Resume Next

'    err = 0
    On Error GoTo errHandle
    Calc_Clique = dbl申领数量
    
    '提取该药品的申领阀值，为零则直接退出
    gstrSQL = "Select Nvl(申领阀值,0) AS 阀值 From 药品规格 Where 药品ID=[1]"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取该药品的申领阀值", lng药品id)

    If err <> 0 Then Exit Function
    If rstemp!阀值 = 0 Then Exit Function
    dbl阀值 = rstemp!阀值
    
    '算法(余数与阀值的一半进行比较（四舍五入），如果小于，舍掉，大于则进位
    dbl商 = Int(dbl申领数量 / dbl阀值)
    dbl余 = dbl申领数量 - (dbl阀值 * dbl商)
    If dbl余 >= (dbl阀值 / 2) And Not bln校验 Then
        dbl商 = dbl商 + 1
    End If
    
    dbl实际数量 = dbl商 * dbl阀值
    Calc_Clique = dbl实际数量
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub Logogram(ByVal staVal As StatusBar, ByVal bytType As Byte)
'简码方式
'staVal: StartusBar控件
'bytType: 0=拼音; 1=五笔;  当前简码状态
    Dim i As Integer
    For i = 1 To staVal.Panels.count
        If staVal.Panels(i).Key = "PY" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrInset
                zldatabase.SetPara "简码方式", 0
                gint简码方式 = 0
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrRaised
            End If
        ElseIf staVal.Panels(i).Key = "WB" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrRaised
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrInset
                zldatabase.SetPara "简码方式", 1
                gint简码方式 = 1
            End If
        End If
    Next
End Sub

Public Function GetDeptStationNode(ByVal lngDeptId As Long) As String
'获取部门所属站点信息
    Dim rsSQL As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle
    strTmp = "select 站点 from 部门表 where id=[1]"
    Set rsSQL = zldatabase.OpenSQLRecord(strTmp, "获取部门所属站点信息", lngDeptId)
    If Not rsSQL.EOF Then
        GetDeptStationNode = Nvl(rsSQL!站点)
    End If
    rsSQL.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '功能:将集点移动控件中:2008-07-08 16:48:35
    err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    objCtl.SetFocus
End Sub

Public Function Select部门选择器(ByVal FrmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str工作性质 As String = "", _
    Optional bln操作员 As Boolean = False, _
    Optional ByVal str服务对象 As String, _
    Optional strsql As String = "") As Boolean
    '------------------------------------------------------------------------------
    '功能:部门选择器
    '参数:objCtl-指定控件
    '     strSearch-要搜索的条件
    '     str工作性质-工作性质:如"V,W,K"
    '     bln操作员-是否加操作员限制
    '     strSQL-直接根据SQL获取数据(但部门表的别名一定要是A)
    '返回:成功,返回true,否则返回False
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim vRect As RECT
    Dim rstemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    On Error GoTo errHandle
    strTittle = "部门选择器"
    vRect = zlControl.GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    
    strKey = GetMatchingSting(strSearch, False)
    
    If strsql <> "" Then
    
        gstrSQL = strsql
    Else
        gstrSQL = "" & _
        "   Select /*+ Rule*/ distinct a.Id,a.上级id,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
        "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间"
    
        If str工作性质 = "" And bln操作员 = False Then
            gstrSQL = gstrSQL & vbCrLf & _
            "   From 部门表 a" & _
            "   Where 1=1"
        Else
            gstrSQL = gstrSQL & vbCrLf & _
            "   From 部门表 a, 部门性质分类 b,部门性质说明 c," & _
            IIf(str工作性质 = "", "", "       (Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) J") & _
            "   Where c.工作性质 = b.名称" & IIf(str工作性质 = "", "(+)", " and B.编码=J.column_value ") & _
            "         AND a.id = c.部门id and c.服务对象 in (select * from Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)))" & _
            IIf(bln操作员 = False, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
            "   and  (a.撤档时间>=to_date('3000-01-01','yyyy-mm-dd') or a.撤档时间 is null ) And (a.站点=[5] or a.站点 is null) "
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.编码 like upper([3]) or a.简码 like upper([3]) or a.名称 like [3] )"
        If IsNumeric(strSearch) Then                         '如果是数字,则只取编码
            If Mid(gtype_UserSysParms.Para_输入方式, 1, 1) = "1" Then strFind = " And (A.编码 Like Upper([3]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.输入全是字母时只匹配简码
            '0-拼音码,1-五笔码,2-两者
            '.int简码方式 = Val(zlDatabase.GetPara("简码方式" ))
            If Mid(gtype_UserSysParms.Para_输入方式, 2, 1) = "1" Then strFind = " And  (a.简码 Like Upper([3]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  '全汉字
            strFind = " And a.名称 Like [3] "
        End If
    End If
    
    If strSearch = "" And str工作性质 = "" And bln操作员 = False And strsql = "" Then
        gstrSQL = gstrSQL & _
        "   Start With A.上级id Is Null Connect By Prior A.ID = A.上级id "
    Else
        gstrSQL = gstrSQL & vbCrLf & strFind & vbCrLf & " Order by A.编码"
    End If
    
    If strSearch = "" And str工作性质 = "" And bln操作员 = False And strsql = "" Then
        '分上下级
        Set rstemp = zldatabase.ShowSQLSelect(FrmMain, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey, str服务对象)
    Else
        Set rstemp = zldatabase.ShowSQLSelect(FrmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, glngUserId, str工作性质, strKey, str服务对象, gstrNodeNo)
    End If
    If blnCancel = True Then
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    If rstemp Is Nothing Then
        MsgBox "没有满足条件的部门,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    Call zlCtlSetFocus(objCtl, True)
    If UCase(TypeName(objCtl)) = UCase("ComboBox") Then
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rstemp!Id) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            MsgBox "你选择的部门在下拉列表中不存在,请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    Else
        objCtl.Text = Nvl(rstemp!编码) & "-" & Nvl(rstemp!名称)
        objCtl.Tag = Val(rstemp!Id)
    End If
    OS.PressKey vbKeyTab
    Select部门选择器 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '     blnUpper-是否转换在大写
    '返回:返回加匹配串%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Sub SetTip(ByVal objControl As Object, ByVal strTip As String)
    '功能:根据显示指定的控件的提示文本
    'objControl:需要提示的控件
    'strTip:已组织好的提示文本
    Call zlCommFun.ShowTipInfo(objControl.hWnd, strTip, True, True, 8800)
End Sub

Public Sub SetSelectorRS( _
    ByVal byt编辑模式 As Byte, _
    ByVal strModeName As String, _
    Optional ByVal lng来源库房 As Long = 0, _
    Optional ByVal lng目标库房 As Long = 0, _
    Optional ByVal lng使用部门 As Long = 0, _
    Optional ByVal lng供应商 As Long = 0, _
    Optional ByVal byt领用方式 As Byte = 0, _
    Optional ByVal bln包含停用药品 As Boolean = False, _
    Optional ByVal bln盘无存储库房药品 As Boolean = False, _
    Optional ByVal byt盘点单据 As Byte = 0, _
    Optional ByVal bln检测库存 As Boolean = True _
    )
'----------------------------------------------------------------------------------------
'功能：初始化grsMaster、grsMasterInput、grsSlave对象，
'      为调用药品选择器(frmSelector)做数据准备。
'参数：
'  byt编辑模式： 1：入库； 2：出库
'  lng来源库房：
'----------------------------------------------------------------------------------------
    Const CON_FMT = "'999999999990.99999'"
    
    Dim strsql As String, strTmp As String
    Dim strUnit As String, strConversionUnit As String
    Dim rstemp As ADODB.Recordset
    Dim IntStockCheck As Integer
    Dim intUnit As Integer, intCostDigit As Integer, intPricedigit As Integer, intNumberDigit As Integer, intMoneyDigit As Integer
    
    On Error GoTo errHandle
    With grsMaster
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    With grsMasterInput
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    With grsSlave
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    
    '数量单位
    If strModeName = "药品申领管理" Or strModeName = "药品移库管理" Then
        Call GetDrugDigit(lng使用部门, strModeName, intUnit, intCostDigit, intPricedigit, intNumberDigit, intMoneyDigit)
    Else
        Call GetDrugDigit(IIf(lng来源库房 = 0, lng目标库房, lng来源库房), strModeName, intUnit, intCostDigit, intPricedigit, intNumberDigit, intMoneyDigit)
    End If
    Select Case intUnit
        Case 1: strConversionUnit = "1"
        Case 2: strConversionUnit = "d.门诊包装"
        Case 3: strConversionUnit = "d.住院包装"
        Case Else
            strConversionUnit = "d.药库包装"
    End Select
    
    '检查库存
    If bln检测库存 = True And strModeName = "药品申领管理" Then
        bln检测库存 = (Val(zldatabase.GetPara("药品按批次出库", glngSys, 1343, 0)) = 1)
    End If
    
    '检查并执行调价
    Call AutoAdjustPrice_Batch
    
    '提取库存检查参数，确定库存不足的不提取数据
    strsql = "Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1] "
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "获取是否库存检查设置", lng来源库房)
    If Not rstemp.EOF Then IntStockCheck = Nvl(rstemp!库存检查, 0)
    rstemp.Close
    
    '*选择模式的数据集*'
    strsql = _
        "Select " & _
        " d.剂型,d.中药形态, d.药名编码, d.通用名称, d.药品来源 As 来源, d.基本药物, d.药典id, d.用途分类id, d.剂量单位, d.药品编码, d.药品名称, " & _
        " d.商品名, d.规格, d.产地 As 生产商, Decode(s.原产地, Null, d.原产地, s.原产地) as 原产地, d.药名id, d.药品id," & _
        " trim(to_char(d.初始成本价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) 上次采购价, " & _
        " trim(to_char(Decode(d.时价, '是', Decode(s.平均售价, Null, Nvl(d.上次售价,p.售价), s.平均售价), p.售价) * " & strConversionUnit & ", '99999999999990." & String(intPricedigit, "0") & "')) 售价, " & _
        " d.售价单位, d.剂量系数 As 售价包装," & _
        " d.门诊单位, d.门诊包装, d.住院单位, d.住院包装, d.药库单位, d.药库包装, " & _
        " nvl(trim(to_char(s.可用数量 / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "')),0) 可用数量, " & _
        " s.库存数量,s.库存金额,s.库存差价,d.最大效期 有效期, d.药库分批, d.药房分批, d.时价," & _
        " trim(to_char(d.指导批发价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as 指导批发价, " & _
        " d.加成率, e.库房货位, d.批准文号, s.库存数量 实际数量, " & _
        " s.留存数量, d.合同单位, d.药价级别,e.领用标志,d.停用,d.上次供应商 " & vbNewLine & _
        "From " & vbNewLine & _
        "  (SELECT DISTINCT J.名称 剂型,Decode(c.类别, '7', Decode(d.中药形态, 1, '饮片', 2, '免煎剂', '散装'), '') As 中药形态,A.名称 商品名, C.编码 药名编码,C.名称 通用名称, 0 AS 药典ID,C.编码 药品编码,C.名称 药品名称," & vbNewLine & _
        "     C.规格,C.产地,d.原产地,C.类别,C.计算单位 AS 售价单位,DECODE(C.是否变价,1,'是','否') 时价,D.药品来源,D.基本药物,D.批准文号, D.药名ID," & vbNewLine & _
        "     D.药品ID, nvl(to_char(D.最大效期,'9999990'),0) 最大效期," & vbNewLine & _
        "     DECODE(D.药库分批,1,'是','否') 药库分批,DECODE(D.药房分批,1,'是','否') 药房分批," & vbNewLine & _
        "     to_char(D.剂量系数, " & CON_FMT & ") 剂量系数," & vbLf & _
        "     D.门诊单位, to_char(D.门诊包装, " & CON_FMT & ") 门诊包装," & vbNewLine & _
        "     D.住院单位, to_char(D.住院包装, " & CON_FMT & ") 住院包装," & vbNewLine & _
        "     D.药库单位, to_char(D.药库包装, " & CON_FMT & ") 药库包装," & vbNewLine & _
        "     D.指导批发价, nvl(D.成本价,0) 初始成本价,D.加成率,D.药价级别," & vbNewLine & _
        "     M.分类ID AS 用途分类ID,M.计算单位 AS 剂量单位,Q.名称 As 合同单位,Decode(Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')), To_Date('3000-01-01', 'yyyy-mm-dd'), '否','是') As 停用,d.上次售价,f.名称 上次供应商  " & vbNewLine
    strsql = strsql & _
        "   FROM 收费项目目录 C,药品规格 D,收费项目别名 A,药品剂型 J,药品特性 T,诊疗项目目录 M,供应商 Q, 诊疗分类目录 E,供应商 F " & vbNewLine & _
        IIf(lng来源库房 <> 0, "     ,(Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[2] Group By 执行科室ID,收费细目ID) K", "") & vbNewLine & _
        IIf(lng目标库房 <> 0, "     ,(Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[3] Group By 执行科室ID,收费细目ID) I ", "") & vbNewLine & _
        "   WHERE C.ID=D.药品ID AND D.药名ID=T.药名ID AND T.药名ID=M.ID and m.分类id=e.id AND M.类别 IN ('5','6','7') " & _
        IIf(lng来源库房 <> 0, "     And D.药品ID=K.收费细目ID" & IIf(bln盘无存储库房药品, "(+)", ""), "") & _
        IIf(lng目标库房 <> 0, "     And D.药品ID=I.收费细目ID" & IIf(bln盘无存储库房药品, "(+)", ""), "") & _
        "     AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 " & _
        "     And (C.站点 = [1] or c.站点 is null) AND T.药品剂型=J.名称(+) And D.合同单位ID=Q.ID(+) And D.上次供应商ID=f.ID(+) " & _
        IIf(bln包含停用药品 = False, " And (C.撤档时间 Is Null Or To_char(C.撤档时间,'yyyy-MM-dd')='3000-01-01') ) D,", ") D,") & vbNewLine & _
        "(Select 收费细目id, 现价 售价 " & _
        " From 收费价目 Where (Sysdate Between 执行日期 And 终止日期 or Sysdate>=执行日期 And 终止日期 Is Null)" & _
        GetPriceClassString("") & ") P," & vbNewLine
    If byt领用方式 = 1 Then
       '向留存领药
       strsql = strsql & _
           "(Select a.药品id,Max(上次产地) AS 产地,max(a.原产地) as 原产地,Sum(a.可用数量) 可用数量," & _
           " To_Char(Sum(a.实际数量), " & CON_FMT & ") 库存数量," & _
           " To_Char(Sum(a.实际金额), " & CON_FMT & ") 库存金额," & _
           " To_Char(Sum(a.实际差价), " & CON_FMT & ") 库存差价," & _
           " Decode(Sum(nvl(实际数量,0)), 0, null, Sum(a.实际金额) / Sum(a.实际数量)) As 平均售价," & _
           " To_Char(Sum(b.实际数量), '99999999999990.99') 留存数量 " & vbNewLine & _
           "From 药品库存 A, 药品留存 B " & vbNewLine & _
           "Where a.性质=1 and a.药品id=b.药品id And a.库房id=b.库房id and b.科室id=[3] and b.期间=to_date(sysdate,'yyyy') "
    Else
       '向药房领药
       strsql = strsql & _
           "(Select a.药品id, Max(a.上次产地) AS 产地, max(a.原产地) as 原产地,Sum(a.可用数量) 可用数量," & _
           " Sum(a.实际数量) 库存数量," & _
           " Sum(a.实际金额) 库存金额," & _
           " Sum(a.实际差价) 库存差价," & _
           " Decode(Sum(nvl(实际数量,0)), 0, null, Sum(a.实际金额) / Sum(a.实际数量)) As 平均售价," & _
           " '' 留存数量 " & vbNewLine & _
           "From 药品库存 A " & vbNewLine & _
           "Where 性质=1 "
    End If
    If lng来源库房 <> 0 Or lng目标库房 <> 0 Then
       strsql = strsql & " And a.库房ID=" & IIf(lng来源库房 = 0, "[3]", "[2]")
    End If
    strsql = strsql & vbNewLine & _
       "Group By a.药品id) S," & vbNewLine & _
       "(Select 药品ID,库房ID,库房货位,领用标志 From 药品储备限额 Where 库房ID=[2]) E " & vbNewLine & _
       "Where D.药品ID=P.收费细目ID And D.药品ID=S.药品ID" & IIf(Not (IntStockCheck = 2 And byt编辑模式 = 2) Or byt盘点单据 = 1 Or Not bln检测库存, "(+)", "") & _
       "  And D.药品ID=E.药品ID(+) " & vbNewLine & _
       "Order By D.药名编码,D.药品编码 "
    Set grsMaster = zldatabase.OpenSQLRecord(strsql, "药品规格", gstrNodeNo, lng来源库房, lng目标库房)
    
    
    '*录入模式的数据集*'
    strsql = _
        "Select " & _
        " d.剂型, d.药名编码, d.通用名称, d.药品来源 来源, d.基本药物, d.药典id, d.用途分类id, d.剂量单位, d.药品编码, f.名称 药品名称, " & _
        " d.商品名, d.规格, d.产地 As 生产商, Decode(s.原产地, Null, d.原产地, s.原产地) as 原产地, d.药名id, d.药品id," & _
        " trim(to_char(d.初始成本价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) 上次采购价, " & _
        " trim(to_char(Decode(d.时价, '是', Decode(s.平均售价, Null, p.售价, s.平均售价), p.售价) * " & strConversionUnit & ", '99999999999990." & String(intPricedigit, "0") & "')) 售价, " & _
        " d.售价单位, d.剂量系数 售价包装,d.门诊单位, d.门诊包装, d.住院单位, d.住院包装, d.药库单位, d.药库包装, " & _
        " nvl(trim(to_char(s.可用数量 / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "')),0) 可用数量, " & _
        " s.库存数量,s.库存金额,s.库存差价, " & _
        " d.最大效期 有效期, d.药库分批, d.药房分批, d.时价," & _
        " trim(to_char(d.指导批发价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as 指导批发价, " & _
        " trim(to_char(d.指导零售价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as 指导零售价, " & _
        "d.加成率, e.库房货位, d.批准文号, s.库存数量 实际数量," & _
        " s.留存数量, d.合同单位, d.药价级别,e.领用标志, Max(Decode(码类, '1', 简码, Null)) 简码, Max(Decode(码类, '3', 简码, Null)) 数字简码,Max(Decode(码类, '2', 简码, Null)) 五笔码,d.停用,d.上次供应商 " & vbNewLine & _
        "From " & vbNewLine & _
        "  (SELECT DISTINCT J.名称 剂型,C.编码 药名编码,C.名称 AS 通用名称,0 AS 药典ID,M.分类ID AS 用途分类ID,M.计算单位 AS 剂量单位, " & _
        "   C.编码 AS 药品编码, a.名称 As 商品名, c.规格, c.产地, d.原产地, d.药品来源, d.基本药物, d.批准文号, d.药名id, " & _
        "   d.药品id, c.计算单位 As 售价单位, nvl(to_char(d.最大效期, '9999990'),0) 最大效期, " & _
        "   DECODE(D.药库分批,1,'是','否') 药库分批, DECODE(D.药房分批,1,'是','否') 药房分批, " & _
        "   to_char(D.剂量系数, " & CON_FMT & ") 剂量系数," & vbLf & _
        "   D.门诊单位, to_char(D.门诊包装, " & CON_FMT & ") 门诊包装," & vbNewLine & _
        "   D.住院单位, to_char(D.住院包装, " & CON_FMT & ") 住院包装," & vbNewLine & _
        "   D.药库单位, to_char(D.药库包装, " & CON_FMT & ") 药库包装," & vbNewLine & _
        "   D.指导批发价,d.指导零售价,nvl(D.成本价,0) 初始成本价, D.加成率, q.名称 合同单位, D.药价级别, " & _
        "   DECODE(C.是否变价,1,'是','否') 时价,Decode(Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')), To_Date('3000-01-01', 'yyyy-mm-dd'), '否','是') As 停用,f.名称 上次供应商 " & vbNewLine
    strsql = strsql & _
        "   From  收费项目目录 C,药品规格 D,收费项目别名 A,药品剂型 J,药品特性 T,诊疗项目目录 M,供应商 Q,供应商 F" & vbNewLine & _
        IIf(lng来源库房 <> 0, "     ,(Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[2] Group By 执行科室ID,收费细目ID) K", "") & vbNewLine & _
        IIf(lng目标库房 <> 0, "     ,(Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[3] Group By 执行科室ID,收费细目ID) I ", "") & vbNewLine & _
        "   Where c.Id = d.药品id And d.药名id = t.药名id And t.药名id = m.Id And m.类别 In ('5', '6', '7') And d.药品id = a.收费细目id(+) " & _
        "     And a.性质(+) = 3 And t.药品剂型 = j.名称(+) And d.合同单位id = q.Id(+) And D.上次供应商ID=f.ID(+) " & _
        IIf(lng来源库房 <> 0, "     And D.药品ID=K.收费细目ID" & IIf(bln盘无存储库房药品, "(+)", ""), "") & _
        IIf(lng目标库房 <> 0, "     And D.药品ID=I.收费细目ID" & IIf(bln盘无存储库房药品, "(+)", ""), "") & _
        IIf(bln包含停用药品 = False, " And (C.撤档时间 Is Null Or To_char(C.撤档时间,'yyyy-MM-dd')='3000-01-01') ) D,", ") D,") & vbNewLine & _
        "  (Select 收费细目id, Trim(To_Char(现价, '999999999990." & String(7, "0") & "')) 售价 " & _
        "   From 收费价目 Where (Sysdate Between 执行日期 And 终止日期 or Sysdate>=执行日期 And 终止日期 Is Null)" & _
        GetPriceClassString("") & ") P," & vbNewLine

    If byt领用方式 = 1 Then
       '向留存领药
       strsql = strsql & _
           "(Select a.药品id,Max(上次产地) AS 产地, max(a.原产地) as 原产地,Sum(a.可用数量) 可用数量," & _
           " To_Char(Sum(a.实际数量), " & CON_FMT & ") 库存数量," & _
           " To_Char(Sum(a.实际金额), " & CON_FMT & ") 库存金额," & _
           " To_Char(Sum(a.实际差价), " & CON_FMT & ") 库存差价," & _
           " Decode(Sum(Nvl(实际数量, 0)), 0, Null, Sum(a.实际金额) / Sum(a.实际数量)) As 平均售价, " & _
           " To_Char(Sum(b.实际数量), '99999999999990.99') 留存数量 " & vbNewLine & _
           "From 药品库存 A, 药品留存 B " & vbNewLine & _
           "Where a.性质=1 and a.药品id=b.药品id And a.库房id=b.库房id and b.科室id=[3] and b.期间=to_date(sysdate,'yyyy') "
    Else
       '向药房领药
       strsql = strsql & _
           "(Select a.药品id, Max(a.上次产地) AS 产地, max(a.原产地) as 原产地,Sum(a.可用数量) 可用数量," & _
           " To_Char(Sum(a.实际数量), " & CON_FMT & ") 库存数量," & _
           " To_Char(Sum(a.实际金额), " & CON_FMT & ") 库存金额," & _
           " To_Char(Sum(a.实际差价), " & CON_FMT & ") 库存差价," & _
           " Decode(Sum(Nvl(实际数量, 0)), 0, Null, Sum(a.实际金额) / Sum(a.实际数量)) As 平均售价, " & _
           " '' 留存数量 " & vbNewLine & _
           "From 药品库存 A " & vbNewLine & _
           "Where 性质=1 "
    End If
    If lng来源库房 <> 0 Or lng目标库房 <> 0 Then
       strsql = strsql & " And a.库房ID=" & IIf(lng来源库房 = 0, "[3]", "[2]")
    End If
    strsql = strsql & vbNewLine & _
       "Group By a.药品id) S," & vbNewLine & _
       "(Select 药品ID,库房ID,库房货位,领用标志 From 药品储备限额 Where 库房ID=" & IIf(byt编辑模式 = 2, "[2]", "[3]") & ") E, 收费项目别名 F " & vbNewLine & _
       "Where D.药品ID=P.收费细目ID And D.药品ID=S.药品ID" & IIf(Not (IntStockCheck = 2 And byt编辑模式 = 2) Or byt盘点单据 = 1 Or Not bln检测库存, "(+)", "") & _
       "  And D.药品ID=E.药品ID(+) And d.药品id = f.收费细目id(+) " & vbNewLine & _
       "Group By d.剂型, d.药名编码, d.通用名称, d.药品来源, d.基本药物, d.药典id, d.用途分类id, d.剂量单位, d.药品编码, f.名称, d.商品名, d.规格, d.产地," & vbNewLine & _
       "  Decode(s.原产地, Null, d.原产地, s.原产地), d.药名id, d.药品id, " & vbNewLine & _
       "  trim(to_char(d.初始成本价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')), " & vbNewLine & _
       "  trim(to_char(Decode(d.时价, '是', Decode(s.平均售价, Null, p.售价, s.平均售价), p.售价) * " & strConversionUnit & ", '99999999999990." & String(intPricedigit, "0") & "'))," & vbNewLine & _
       "   d.售价单位,d.剂量系数, d.门诊单位, d.门诊包装, d.住院单位, d.住院包装, d.药库单位, d.药库包装," & vbNewLine & _
       "  nvl(trim(to_char(s.可用数量 / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "')),0), s.库存数量, s.库存金额, s.库存差价, " & vbNewLine & _
       "  d.最大效期, d.药库分批, d.药房分批, d.时价,trim(to_char(d.指导批发价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')), " & vbNewLine & _
       "  trim(to_char(d.指导零售价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')), d.加成率,e.库房货位, d.批准文号, s.库存数量," & vbNewLine & _
       "  s.留存数量, d.合同单位, d.药价级别, e.领用标志, d.停用, d.上次供应商 " & vbNewLine & _
       "Order By D.药名编码,D.药品编码 "
    Set grsMasterInput = zldatabase.OpenSQLRecord(strsql, "药品规格", gstrNodeNo, lng来源库房, lng目标库房, IIf(gint简码方式 = 0, 1, 2))
    
    '*药品分批*'
    If byt编辑模式 = 2 Then
        strsql = _
            "Select Rid,库房,药品ID,批次,入库日期,批号,生产日期,产地 as 生产商,原产地,成本价,售价,时价,门诊单位,门诊包装,住院单位,住院包装,药库单位,药库包装," & _
            "  有效期,实际数量,可用数量,库存数量,库存金额,库存差价,上次供应商ID,批准文号,供应商 " & vbLf & _
            "From (Select Distinct 2 Rid, p.名称 库房, k.药品id, nvl(k.批次,0) 批次, To_Char(b.入库日期, 'YYYY-MM-DD') As 入库日期, k.上次批号 批号," & _
            "  To_Char(k.上次生产日期, 'YYYY-MM-DD') 生产日期, k.上次产地 产地, Decode(k.原产地, Null, d.原产地, k.原产地) as 原产地, k.平均成本价 成本价," & _
            "  Decode(Nvl(k.批次, 0), 0, Decode(Sign(k.实际数量), 1, k.实际金额 / decode(nvl(k.实际数量,0), 0, 1, k.实际数量), A.现价) " & _
            "        ,Nvl(k.零售价, k.实际金额 / decode(nvl(k.实际数量,0), 0, 1, k.实际数量) ) ) 售价," & _
            "  Nvl(k.零售价, k.实际金额 / decode(nvl(k.实际数量,0), 0, 1, k.实际数量) ) 时价," & _
            "  D.门诊单位, to_char(D.门诊包装, " & CON_FMT & ") 门诊包装," & _
            "  D.住院单位, to_char(D.住院包装, " & CON_FMT & ") 住院包装," & _
            "  D.药库单位, to_char(D.药库包装, " & CON_FMT & ") 药库包装," & _
            "  k.效期" & IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "-1", "") & " 有效期," & _
            "  k.实际数量, k.可用数量, k.实际数量 库存数量, k.实际金额 库存金额, k.实际差价 库存差价, k.上次供应商id, k.批准文号,f.名称 供应商 " & vbNewLine & _
            "From 部门表 P, 药品规格 D, 药品库存 K, 药品入库信息 B, 收费价目 A,供应商 F " & vbNewLine & _
            "Where k.库房id = p.Id And d.药品id = k.药品id And d.药品id=a.收费细目id " & GetPriceClassString("A") & _
            "  And k.性质 = 1 And k.药品id = b.药品id(+) And k.库房id = b.库房id(+) And nvl(k.批次,0)  = nvl(b.批次(+),0) And k.库房id = [1] And K.上次供应商ID=f.ID(+) "
        If byt盘点单据 = 1 Then
            strsql = strsql & " And (K.实际数量<>0 Or K.实际金额<>0 Or K.实际差价<>0) ) " & vbNewLine
'        ElseIf byt盘点单据 = 2 Then
'            '1303 如果是库存差价调整模块，则允许过滤库存数量为0的药品记录
'            gstrSQL = strSQL & " ) " & vbNewLine
        Else
            strsql = strsql & " And K.实际数量<>0 ) " & vbNewLine
        End If
        If gtype_UserSysParms.P150_药品出库优先算法 = 0 Then
            strsql = strsql & "Order By 药品id, 批次 "
        Else
            strsql = strsql & "Order By 药品id, 有效期, 批次 "
        End If

        Set grsSlave = zldatabase.OpenSQLRecord(strsql, "药品分批", IIf(lng来源库房 = 0, lng目标库房, lng来源库房))
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ReleaseSelectorRS()
    If Not grsMaster Is Nothing Then
        If grsMaster.State = adStateOpen Then grsMaster.Close
        Set grsMaster = Nothing
    End If
    
    If Not grsMasterInput Is Nothing Then
        If grsMasterInput.State = adStateOpen Then grsMasterInput.Close
        Set grsMasterInput = Nothing
    End If
    
    If Not grsSlave Is Nothing Then
        If grsSlave.State = adStateOpen Then grsSlave.Close
        Set grsSlave = Nothing
    End If
End Sub

Public Function GetVSFlexRows(ByVal vsfVal As VSFlexGrid, Optional ByVal blnHidden = False) As Long
'--------------------------------------------------------------
'功能：求VSFlexGrid的行数量，含列头行
'参数：
'  blnHidden：True计算非隐藏的行数；False计算隐藏的行数。
'返回：行数量
'--------------------------------------------------------------
    Dim i As Long, lngRows As Long
    For i = 0 To vsfVal.rows - 1
        If blnHidden Then
            If vsfVal.RowHidden(i) Then lngRows = lngRows + 1
        Else
            If vsfVal.RowHidden(i) = False Then lngRows = lngRows + 1
        End If
    Next
    GetVSFlexRows = lngRows
End Function


Public Sub GetPriceClass()
    '根据登录站点获取药品的价格等级
    Dim rsData As ADODB.Recordset
    
    If gstrNodeNo <> "" And gstrNodeNo <> "-" Then
        gstrSQL = " Select a.价格等级 " & _
            " From 收费价格等级应用 A, 收费价格等级 B " & _
            " Where a.价格等级 = b.名称 And a.性质 = 0 And b.是否适用药品 = 1 And a.站点 = [1] And Nvl(b.撤档时间, Sysdate + 1) > Sysdate "
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "GetPriceClass", gstrNodeNo)
        
        If rsData.RecordCount > 0 Then gstrPriceClass = rsData!价格等级
    End If
End Sub


Public Function GetPriceClassString(strTableName As String) As String
    '根据传入表的别名返回价格等级的条件串
    GetPriceClassString = " And " & IIf(strTableName = "", "价格等级 Is Null ", strTableName & ".价格等级 Is Null ")
    
End Function
