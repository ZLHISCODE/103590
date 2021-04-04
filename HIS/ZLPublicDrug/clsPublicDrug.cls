VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    If zlGetComLib = False Then Exit Function
    
    Call GetUserInfo
'    Call GetSysParms
'    Call GetStockCheckRule
    
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCheckPriceAdjustBySell(ByVal lng药品ID As Long, ByVal lng药房id As Long, Optional ByVal blnShowMsg As Boolean = True) As Boolean
    '零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
    '定价药品：售价是固定的，比较所有药房的成本价，如果存在不一致的就不能销售出库
    '时价药品：比较药房库存记录的零售价和成本价，如果存在不一致的就不能销售出库
    '销售出库时只判断药房
    '参数：lng药品id-药品规格ID，lng药房id-销售出库的药房ID，blnShowMsg-是否显示提示(默认为显示)
    '返回：True-正常进行销售出库；false-不能进行销售出库
    
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '如果没开启全局的零差价管理，则不进行后续检查，返回true
    If Val(gobjDatabase.GetPara(275, 100, , 0)) = 0 Then zlCheckPriceAdjustBySell = True: Exit Function
    
    '检查有无库存
    If lng药品ID > 0 Then
        If lng药房id > 0 Then
            gstrSQL = "Select 1 from 药品库存 Where 性质=1 and 药品id=[1] and 库房id=[2] " & _
                " And Not (批次 = 0 And 可用数量 < 0 And 实际数量 = 0 And 实际金额 = 0 And 实际差价 = 0)"
        Else
            gstrSQL = "Select 1 from 药品库存 Where 性质=1 and 药品id=[1] " & _
                " And Not (批次 = 0 And 可用数量 < 0 And 实际数量 = 0 And 实际金额 = 0 And 实际差价 = 0)"
        End If
        Set rsData = gobjDatabase.OpenSQLRecord(gstrSQL, "zlCheckPriceAdjustBySell", lng药品ID, lng药房id)
        
        If rsData.EOF Then
            '无库存时，从收费价目取售价，从药品规格取成本价，并比较价格
            gstrSQL = "Select a.成本价, b.现价 As 售价 " & _
                " From 药品规格 A, 收费价目 B " & _
                " Where a.药品id = b.收费细目id And (Sysdate Between b.执行日期 And b.终止日期) And Nvl(a.是否零差价管理, 0) = 1 " & _
                " And b.现价 <> a.成本价 And a.药品id = [1] "
            Set rsData = gobjDatabase.OpenSQLRecord(gstrSQL, "zlCheckPriceAdjustBySell", lng药品ID)
            
            If rsData.EOF Then
                '没找到表示价格一致或未启用零差价管理
                zlCheckPriceAdjustBySell = True
            Else
                '找到表示价格不一致
                zlCheckPriceAdjustBySell = False
            End If
            
            Exit Function
        End If
    End If
    

    gstrSQL = "Select 药品id, 通用名, 规格, 0 As 库房id, '' As 库房, 生产商, '' As 批号, 批次, 单位, 药库包装, 售价, Sum(成本价 * 实际数量) / Sum(实际数量) As 成本价, 是否时价" & vbNewLine & _
        " From (Select a.药品id, '['|| c.编码 || ']'|| c.名称||decode(c.产地,null,null,'('||c.产地||')') ||c.规格 As 通用名, c.规格, c.产地 As 生产商, Null As 批次, a.药库单位 As 单位, a.药库包装, b.现价 As 售价," & vbNewLine & _
        "              d.平均成本价 As 成本价, 0 As 是否时价, d.实际数量" & vbNewLine & _
        "       From 药品规格 A, 收费价目 B, 收费项目目录 C, 药品库存 D" & vbNewLine & _
        "       Where a.药品id = b.收费细目id And a.药品id = c.Id And a.药品id = d.药品id And d.性质 = 1 And (Sysdate Between b.执行日期 And b.终止日期) And" & vbNewLine & _
        "             (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.是否变价 = 0 And Nvl(a.是否零差价管理, 0) = 1 And" & vbNewLine & _
        "             b.现价 <> d.平均成本价 And d.库房id In(SELECT Distinct 部门id From 部门性质说明  WHERE 工作性质 LIKE '%药房') And a.药品id = [1] And d.库房id = [2] " & vbNewLine & _
        "  And Not (D.批次 = 0 And D.可用数量 < 0 And D.实际数量 = 0 And D.实际金额 = 0 And D.实际差价 = 0))" & vbNewLine & _
        " Group By 药品id, 通用名, 规格, 生产商, 批次, 单位, 药库包装, 售价, 是否时价 " & vbNewLine & _
        " Having Sum(实际数量) <> 0" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select a.药品id, '['|| c.编码 || ']'|| c.名称||decode(c.产地,null,null,'('||c.产地||')') ||c.规格 As 通用名, c.规格, d.库房id, e.名称 As 库房, d.上次产地 As 生产商, d.上次批号 As 批号, d.批次," & vbNewLine & _
        "       a.药库单位 As 单位, a.药库包装, Nvl(d.零售价,0) As 售价, d.平均成本价 As 成本价, 1 As 是否时价" & vbNewLine & _
        " From 药品规格 A, 收费项目目录 C, 药品库存 D, 部门表 E" & vbNewLine & _
        " Where a.药品id = c.Id And a.药品id = d.药品id And d.库房id = e.Id And d.性质 = 1 And c.是否变价 = 1 And" & vbNewLine & _
        "      (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.是否零差价管理, 0) = 1 And Nvl(d.零售价,0) <> d.平均成本价" & vbNewLine & _
        " And d.库房id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房') And a.药品id = [1] And d.库房id = [2] " & vbNewLine & _
        " And Not (d.批次 = 0 And d.可用数量 < 0 And d.实际数量 = 0 And d.实际金额 = 0 And d.实际差价 = 0) " & vbNewLine & _
        " Order By 通用名,库房id,批号"
    Set rsData = gobjDatabase.OpenSQLRecord(gstrSQL, "zlPublicDrug：zlCheckPriceAdjustBySell", lng药品ID, lng药房id)
    
    '没找到不满足零差价管理要求的记录，返回true
    If rsData.EOF Then zlCheckPriceAdjustBySell = True: Exit Function
    
    '找到不满足零差价管理要求的记录时，提示并返回false
    If blnShowMsg = True Then
        MsgBox "在(" & rsData!库房 & ")中药品" & rsData!通用名 & "不满足零差价管理的要求：成本价和售价不一致，不能销售出库。" & vbCrLf & "请联系药房或药剂科进行调价处理。", vbInformation, "zlPublicDrug"
    End If
    
    zlCheckPriceAdjustBySell = False
    
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Private Function zlGetDrugAlllAvailable(ByVal lng药品ID As Long, ByRef dbl可用数量 As Double) As Boolean
    '计算指定药品在所有药房的可用数量，不管库房，批次
    '1.汇总指定药品的可用数量
    '2.排除效期过期的库存药品记录
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select Nvl(Sum(可用数量), 0) As 可用数量 " & _
        " From 药品库存 Where 性质 = 1 And 药品id = [1] And (效期 > Sysdate Or 效期 Is Null) "
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetDrugAlllAvailable", lng药品ID)
    
    dbl可用数量 = rsTemp!可用数量
    
    zlGetDrugAlllAvailable = True
End Function


Private Function zlGetDrugAvailable(ByVal lng库房id As Long, ByVal lng药品ID As Long, ByRef dbl可用数量 As Double) As Boolean
    '计算指定库房，药品的可用数量，不管批次
    '1.汇总指定库房，药品的可用数量
    '2.排除效期过期的库存药品记录
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select Nvl(Sum(可用数量), 0) As 可用数量 " & _
        " From 药品库存 Where 性质 = 1 And 库房id = [1] And 药品id = [2] And (效期 > Sysdate Or 效期 Is Null) "
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetDrugAvailable", lng库房id, lng药品ID)
    
    dbl可用数量 = rsTemp!可用数量
    
    zlGetDrugAvailable = True
End Function


Private Function zlCheckDrugDosage(ByVal byt性质 As Byte, ByVal strNos As String, ByRef strDosageNos As String) As Boolean
    '判断药品是否已经配药
    '药品收发记录中根据传参来批量提取对应的收费，记账类型的发药单据，再判断当中是否有已经配药的药品，返回已配药的单据号
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select /*+ rule*/ a.No " & _
        " From 未发药品记录 A, Table(Cast(f_Str2list([2]) As t_Strlist)) B " & _
        " Where a.单据 = [1] And a.配药人 Is Not Null And a.No = b.Column_Value Order By a.NO "
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetDrugAvailable", IIf(byt性质 = 1, 8, 9), strNos)
    
    strDosageNos = ""
    Do While Not rsTemp.EOF
        strDosageNos = IIf(strDosageNos = "", "", strDosageNos & ",") & rsTemp!NO
        
        rsTemp.MoveNext
    Loop
    
    zlCheckDrugDosage = True
End Function

Private Function zlGetDrugCurrentPrice(ByVal lng库房id As Long, ByVal lng药品ID As Long, ByVal dbl出库数量 As Double, ByRef dbl库存数量 As Double, ByRef dbl出库价格 As Double, ByRef dbl出库金额 As Double) As Boolean
    '获取时价药品售价和应收金额
    '1.根据药品ID和库房ID获取时价药品的售价及出库总金额(应收金额)
    '2.售价的确定:总金额/总数量;如果只是一个批次，直接取药品库存的零售价
    '3.批次选择需要根据出库规则进行处理(即:系统参数:150-药品出库优先算法(0-按批次先进先出，1-按效期最近先出))
    '4.如果当前库存不足，只计算最大允许出库数量的金额
    '5.价格，金额的精度根据系统参数确定(费用精度)
    Dim rsTemp As ADODB.Recordset
    Dim dbl剩余数量 As Double
    Dim dbl实际出库数量 As Double
    Dim dbl当前数量 As Double
    Dim dbl当前价格 As Double
    Dim blnOnlyOneRs As Boolean     '是否只有一条库存记录
    
    dbl库存数量 = 0
    dbl出库价格 = 0
    dbl出库金额 = 0
    
    gstrSQL = "Select Nvl(批次,0) As 批次, 效期, Nvl(实际数量, 0) As 实际数量, Nvl(实际金额, 0) As 实际金额, Nvl(实际差价, 0) As 实际数量, nvl(零售价,0) As 零售价 " & _
        " From 药品库存 Where 性质 = 1 And 库房id = [1] And 药品id = [2] "
    If gtype_UserSysParms.P150_药品出库优先算法 = 1 Then
        gstrSQL = gstrSQL & " Order by 效期,批次 "
    Else
        gstrSQL = gstrSQL & " Order by 批次 "
    End If
        
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetDrugCurrentPrice", lng库房id, lng药品ID)
    
    '实际算法
    '1.按实际数量来计算
    '2.默认是严格控制库存，库存不足按现有最大数量计算
    '3.时价分批药品按“药品库存.零售价”取价格，否则按“实际金额/实际数量”计算价格
    '4.如果实际数量小于0，按0取值，并且价格和金额也为0
    '5.按批次或效期顺序依次分解出库数量
    
    If rsTemp.RecordCount = 0 Then
        zlGetDrugCurrentPrice = True
        Exit Function
    ElseIf rsTemp.RecordCount = 1 Then
        blnOnlyOneRs = True
    End If
        
    dbl剩余数量 = dbl出库数量
    
    Do While Not rsTemp.EOF
        dbl当前价格 = 0
        dbl当前数量 = 0
        
        '时间分批取零售价，否则按库存金额，数量计算
        If rsTemp!批次 > 0 And rsTemp!零售价 > 0 Then
            dbl当前价格 = Round(rsTemp!零售价, gtype_UserSysParms.P157_费用单价保留位数)
        ElseIf rsTemp!实际数量 > 0 Then
            dbl当前价格 = Round(rsTemp!实际金额 / rsTemp!实际数量, gtype_UserSysParms.P157_费用单价保留位数)
        End If
        
        '如果还有剩余数量，本批次也有实际数量可供分配时
        If rsTemp!实际数量 > 0 And dbl剩余数量 > 0 Then
            If dbl剩余数量 - rsTemp!实际数量 <= 0 Then
                '本次分解就够了，没有剩余了
                dbl当前数量 = dbl剩余数量
                dbl剩余数量 = 0
            Else
                '有剩余，需要在下个批次分解
                dbl当前数量 = rsTemp!实际数量
                dbl剩余数量 = dbl剩余数量 - rsTemp!实际数量
            End If
            
            '按本批次分解的出库数量计算本批次的出库金额和并记录累计出库数量
            '出参
            dbl出库金额 = dbl出库金额 + Round(dbl当前价格 * dbl当前数量, gtype_UserSysParms.P9_费用金额保留位数)
            
            dbl实际出库数量 = dbl实际出库数量 + dbl当前数量
        End If
        
        '计算库存实际数量合计
        '出参
        dbl库存数量 = dbl库存数量 + rsTemp!实际数量
    Loop
    
    If blnOnlyOneRs = True Then
        '只有一条库存记录时，直接取前面已计算的价格
        '出参
        dbl出库价格 = dbl当前价格
    ElseIf dbl出库金额 > 0 And dbl实际出库数量 > 0 Then
        '按最后的合计出库金额和出库数量计算出库价格
        '出参
        dbl出库价格 = Round(dbl出库金额 / dbl实际出库数量, gtype_UserSysParms.P157_费用单价保留位数)
    End If
    
    zlGetDrugCurrentPrice = True
    
End Function

Private Function zlGetDrugReturnAmount(ByVal byt性质 As Byte, ByVal strNos As String, ByRef rsData As ADODB.Recordset) As Boolean
    '获取药品发药数据的准退数量
    '根据传入的单据性质和单据号关联药品，费用表查询
    '返回记录集：NO,费用ID,收费细目ID,原始数量，已执行数量，准退数量
    
    gstrSQL = "Select /*+ rule*/" & vbNewLine & _
        " NO, 费用id, 药品id As 收费细目id, Sum(原始数量) As 原始数量, Sum(准退数量) As 已执行数量, Sum(准退数量) As 准退数量" & vbNewLine & _
        "From (Select a.No, a.记录状态, a.序号, a.药品id, a.批次, a.费用id, a.实际数量 * Nvl(a.付数, 1) As 原始数量, 0 As 准退数量" & vbNewLine & _
        "       From 药品收发记录 A, Table(Cast(f_Str2list([2]) As t_Strlist)) B" & vbNewLine & _
        "       Where a.单据 = [1] And a.No = b.Column_Value And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0) And a.审核日期 Is Not Null" & vbNewLine & _
        "       Union All" & vbNewLine & _
        "       Select a.No, a.记录状态, a.序号, a.药品id, a.批次, a.费用id, 0 As 原始数量, a.实际数量 * Nvl(a.付数, 1) As 准退数量" & vbNewLine & _
        "       From 药品收发记录 A, Table(Cast(f_Str2list([2]) As t_Strlist)) B" & vbNewLine & _
        "       Where a.单据 = [1] And a.No = b.Column_Value And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0 Or Mod(a.记录状态, 3) = 2) And" & vbNewLine & _
        "             a.审核日期 Is Not Null)" & vbNewLine & _
        "Group By NO, 费用id, 药品id " & vbNewLine & _
        "Order By NO, 费用id, 药品id "
    Set rsData = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetDrugReturnAmount", IIf(byt性质 = 1, 8, 9), strNos)
    
End Function

Private Function zlTriggerDrugAjustPrice(ByVal str药品IDs As String) As Boolean
    '触发药品调价
    '根据传参依次判断每个药品是否存在已生效的新价格，如果存在则进行调价处理
    Dim arrSql As Variant
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    arrSql = Array()
     
    gstrSQL = "Select Distinct n.Id " & _
               " From 收费项目目录 I, 收费价目 N, 药品规格 P " & _
               " Where i.Id = n.收费细目id And i.Id = p.药品id And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And " & _
                   " n.变动原因 = 0 And Sysdate>n.执行日期 " & _
               " Order By n.id "
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "zlTriggerDrugAjustPrice：查询调价生效记录")
    
    If rsTemp.RecordCount = 0 Then Exit Function
    
    For i = 0 To rsTemp.RecordCount - 1
        gstrSQL = "Zl_药品收发记录_Adjust(" & rsTemp!ID & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
    Next

    For i = 0 To UBound(arrSql)
        Call gobjDatabase.ExecuteProcedure(CStr(arrSql(i)), "zlTriggerDrugAjustPrice：执行调价")
    Next
    
    zlTriggerDrugAjustPrice = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function zlDrugAdjust(lng药品ID As Long) As Boolean
    '功能：药品调价
    '传参：lng药品ID-药品id
    '返回：True-成功,False-失败
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH
    
    gstrSQL = _
        "Select ID From 收费价目 " & _
        " Where Sysdate Between 执行日期 And Nvl(终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
        " And Nvl(变动原因,0)=0 And 收费细目ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "zlDrugAdjust", lng药品ID)
    
    If Not rsTmp.EOF Then
        gstrSQL = "zl_药品收发记录_Adjust(" & rsTmp!ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, App.ProductName)
        rsTmp.MoveNext
    End If
    
    zlDrugAdjust = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlAutoSendDrug(ByVal lng库房id As Long, ByVal byt性质 As Byte, ByVal strNO As String, ByRef strSendSQL As String) As Boolean
    '自动发药
    '根据传参组织发药需要的数据，返回组织好的发药过程SQL语句
    
    strSendSQL = "zl_药品收发记录_处方发药("
    '库房ID
    strSendSQL = strSendSQL & lng库房id
    '单据
    strSendSQL = strSendSQL & "," & IIf(byt性质 = 1, 8, 9)
    'NO
    strSendSQL = strSendSQL & ",'" & strNO & "'"
    '发药人(审核人)
    strSendSQL = strSendSQL & ",'" & UserInfo.用户姓名 & "'"
    '配药人(必须经过配药过程时，则配药人不填)
    strSendSQL = strSendSQL & ",'" & UserInfo.用户姓名 & "'"
    '校验人（开单医生）
    strSendSQL = strSendSQL & ","
    '发药方式
    strSendSQL = strSendSQL & ",1"
    '发药时间
    strSendSQL = strSendSQL & ","
    '操作员编码
    strSendSQL = strSendSQL & ",'" & UserInfo.用户编码 & "'"
    '操作员名称
    strSendSQL = strSendSQL & ",'" & UserInfo.用户姓名 & "'"
    '金额保留位数
    strSendSQL = strSendSQL & "," & gtype_UserSysParms.P9_费用金额保留位数
    '自动审核记账单
    strSendSQL = strSendSQL & ","
    '是否门诊
    strSendSQL = strSendSQL & ",1"
    '核查人
    strSendSQL = strSendSQL & ","
    ''''结束
    strSendSQL = strSendSQL & ")"
    
    zlAutoSendDrug = True
    Exit Function
End Function

Private Function zlGetReceiveNo(ByVal lng科室id As Long, ByVal str开始时间 As String, ByVal str结束时间 As String, ByRef rsDrawNO As ADODB.Recordset) As Boolean
    '获取指定科室的领药号
    '根据传参查询时间范围内的指定领药部门的领药号，并以数据集的形式返回
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 领药号 From 未发药品记录 Where 对方部门id + 0 = [1] And 填制日期 Between [2] And [3] Order By 领药号"
    Set rsDrawNO = gobjDatabase.OpenSQLRecord(gstrSQL, "zlGetReceiveNo", lng科室id, CDate(Format(str开始时间, "yyyy-mm-dd hh:mm:ss")), CDate(Format(str结束时间, "yyyy-mm-dd hh:mm:ss")))
    
    zlGetReceiveNo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

