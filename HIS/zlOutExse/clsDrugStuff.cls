VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************************************************************
'功能：药品及卫生材料相关处理
'   AutoSplitSpeci:针对中草药根据药名（品种）来自动分配药品(自动分解多个规格)
'   BuildNewBillsData:构建新单据数据
'   CheckBill存在非散装草药:检查单据中是否存在非散装草药形态
'   CheckDisable:检查单据中的药品的禁忌情况
'   CheckDrugExist:判断指定药品或卫材(跟踪在用)在单据中是否已经存在
'   CheckDuty:检查指定药品行的职务是否与当前医生的职务相匹配
'   CheckExistRepeatRow:检查分批或时价药品同一药房是否有重复输入
'   CheckExistWindow:发药窗口检查
'   CheckIsputDrug:检查药品是否已进行配药
'   CheckItemUnderSet:检查指定药品/材料在指定库房的库存是否低于储备下限
'   CheckLimit:费用单据药品处方限量检查,适用于记帐单,收费
'   CheckPriceAdjustBySell:零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
'   CheckServiceDept:检查药品的发药药房对应的服务科室
'   CheckStock:库存检查
'   CheckValidity:检查卫生材料的灭菌效期是否过期
'   DrugMachine_Charge:收费或记帐后发药机数据上传
'   DrugMachine_Del:退费或销账后发药机数据上传
'   ExcuteFeeAndRecipeBillDel:执行处方数据作废操作
'   ExcuteSaveNewBills:执行处方数据保存操作
'   ExcuteSaveNewBillsCheck:执行处方保存前检查操作
'   ExecuteDrugBillDel:作废药品处方
'   ExecuteDrugStuffBillDel:执行处方单、发料单的删除
'   ExecuteStuffBillDel:作废卫生材料处方
'   GetDrugStuffBillDelData:获取作废药品、材料的数据，针对历史退费异常
'   GetDrugTotal:获取单据中指定药品在同一药房多行的数量合
'   GetDrugWindowFromBill:从已分配发药窗口的单据中取发药窗口
'   GetOriginalTotal:获取单据中指定药品在同一药房多行的原始数量和
'   GetPrice:获取药品/卫材价格
'   GetStock:获取指定药房指定药品库存(以零售单位)
'   GetStockByDrugName:根据药名,获取指定库存信息
'   GetStockInfo:获取药品在各个药房，药库的库存信息
'   Get处方限量:获取指定药品的处方限量,以零售单位返回
'   Get处方职务:根据药品ID获取其处方职务
'   Get发药窗口:分配发药窗口
'   Get中药形态:获取单据中的中药形态
'   InitCommon:初始化相关基础数据
'   Read发药窗口:根据药房ID获取发药窗口
'   RestWindowVariable:重置发药窗口相关变量
'   ShowStatusCargoSpace:显示库房货位
'   ShowStatusStock:在状态栏显示药品或卫材的库存
'   UpdateNewBillsData:更新新单据的数据

'   *BillDrugDept:读取一张收费单据的药房ID
'   *GetServiceObject:创建费用公共部件
'   *CreateDrugPacker:创建自助发药机(自动化药房)
'   *CreatePublicDrug:动态创建药品公共部件
'   *GetBillNoWindowDrug:获取单据中未分配发药窗口的药品的药房ID
'   *Getcargospace:获取指定药品在指定药房的库房货位
'   *GetDefaultWindow:获取缺省的药房窗口设置
'   *GetFirstRow:获取当前单据中第一个为药品的行号
'   *GetStockCheck:获取药品/卫材的库存检查方式
'   *SaveDrugID:保存当前指定单据序号以前中最后一个含有药品的单据的第一行药品的部门ID
'*************************************************************************************************************************************************

Private mlngModule As Long, mstrPrivs As String
Private mbln处方职务检查 As Boolean '是否进行处方职务检查
Private mbln处方限量检查 As Boolean '是否进行处方限量检查
Private mbln已检查处方限量 As Boolean '处方限量检查, 是否输入时选择了允许

'参数
Private mbyt药品摆药退费方式 As Byte '47400
Private mcolDrugStockCheck As Collection '存放各个药品库房的出库检查方式
Private mcolStuffStockCheck As Collection '存放各个卫材库的出库检查方式
Private Enum StockCheckType       '库存检查方式
    NOCheck = 0                   '不检查
    CheckRemind = 1               '检查提示
    CheckForbid = 2               '检查禁止
End Enum
Private Enum CheckRemindSelect    '库存检查方式为"检查提示"时,操作员所做的选择
    NORemind = 0                  '还未提示
    Allow = 1                     '提示后操作员选择继续
    Ban = 2                       '提示后操作员选择禁止
End Enum

Private mobjPublicDrug As Object
Private mobjService As zlPublicExpense.clsService

Private mobjDrugPacker  As Object ' 自动发药机(更新发药窗口)
Private mobjDrugMachine As Object '自动发药机(新）


Private mcllWindows As Collection 'array(药房ID,窗口)
Private mstr西窗 As String, mstr成窗 As String, mstr中窗 As String '记录门诊病人连续收费的窗口分配
Private mlng西药房 As Long, mlng成药房 As Long, mlng中药房 As Long '记录门诊病人连续收费的药房分配

Public Property Let 西药房窗口(ByVal strNewValue As String)
    mstr西窗 = strNewValue
End Property

Public Property Get 西药房窗口() As String
    西药房窗口 = mstr西窗
End Property

Public Property Let 成药房窗口(ByVal strNewValue As String)
    mstr成窗 = strNewValue
End Property

Public Property Get 成药房窗口() As String
    成药房窗口 = mstr成窗
End Property

Public Property Let 中药房窗口(ByVal strNewValue As String)
    mstr中窗 = strNewValue
End Property

Public Property Get 中药房窗口() As String
    中药房窗口 = mstr中窗
End Property


Public Property Let 西药房(ByVal lngNewValue As Long)
    mlng西药房 = lngNewValue
End Property

Public Property Get 西药房() As Long
    西药房 = mlng西药房
End Property

Public Property Let 成药房(ByVal lngNewValue As Long)
    mlng成药房 = lngNewValue
End Property

Public Property Get 成药房() As Long
    成药房 = mlng成药房
End Property

Public Property Let 中药房(ByVal lngNewValue As Long)
    mlng中药房 = lngNewValue
End Property

Public Property Get 中药房() As Long
    中药房 = mlng中药房
End Property

Public Sub RestWindowVariable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重置发药窗口相关变量
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mcllWindows = Nothing
    mstr西窗 = "": mstr成窗 = "": mstr中窗 = ""
    mlng西药房 = 0: mlng成药房 = 0: mlng中药房 = 0
End Sub

Public Function InitCommon(ByVal lngModule As Long, strPrivs As String, Optional ByVal blnNewBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关基础数据
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    mlngModule = lngModule: mstrPrivs = strPrivs
    If CreatePublicDrug(gcnOracle) = False Then Exit Function
    If GetPubService(mobjService, lngModule) = False Then Exit Function
    
    mbln处方职务检查 = False
    mbln处方限量检查 = False
    mbln已检查处方限量 = False
    
    '参数
    If lngModule = 1121 Or lngModule = 1124 Then
        mbyt药品摆药退费方式 = Val(zlDatabase.GetPara("药品摆药退费方式", glngSys, lngModule, "0"))
    End If
    If (lngModule = 1120 Or lngModule = 1121 Or lngModule = 1122) And blnNewBill Then
        Set mcolDrugStockCheck = GetStockCheck(False)
        Set mcolStuffStockCheck = GetStockCheck(True)
    End If
    
    strSql = "Select '处方职务' As 类别,count(药名ID) As num From 药品特性 Where 处方职务<>'00' Union All " & _
            " Select '处方限量' As 类别,count(药名ID) As num From 药品特性 Where 处方限量>0"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSql, "clsDrugStuff")
    
    rsTmp.Filter = "类别='处方职务'"
    If Not rsTmp.EOF Then mbln处方职务检查 = (rsTmp!Num > 0)
    
    rsTmp.Filter = "类别='处方限量'"
    If Not rsTmp.EOF Then mbln处方限量检查 = (rsTmp!Num > 0)
    
    InitCommon = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreatePublicDrug(cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:动态创建药品公共部件
    '入参:
    '   lngSys-系统号
    '   cnOracle-数据库连接对象
    '   strDBUser-数据库所有者
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjPublicDrug Is Nothing Then CreatePublicDrug = True: Exit Function
    
    On Error Resume Next
    Set mobjPublicDrug = CreateObject("zlPublicDrug.clsPublicDrug")
    If Err <> 0 Then mobjPublicDrug = Nothing
    If mobjPublicDrug Is Nothing Then
        MsgBox "药品公共部件（zlPublicDrug.clsPublicDrug）创建失败，请与系统管员联系！", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo 0
    'Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If mobjPublicDrug.zlInitCommon(glngSys, cnOracle, gstrDBUser) = False Then
        MsgBox "药品公共部件（zlPublicDrug.clsPublicDrug）初始化失败，请与系统管员联系！", vbInformation, gstrSysName
        Set mobjPublicDrug = Nothing: Exit Function
    End If
    CreatePublicDrug = True
End Function

Public Function GetPubService(ByRef objPubService As Object, ByVal lngMoudle As Long) As Boolean
    '创建服务公共部件
    If gobjPubService Is Nothing Then
        Set gobjPubService = CreateObject("zlPublicExpense.clsService")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense.clsService)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    If gobjPubService Is Nothing Then Exit Function
    If gobjPubService.zlInitCommon(glngSys, lngMoudle, gcnOracle, gstrDBUser) = False Then Exit Function
    
    Set objPubService = gobjPubService
    GetPubService = True
End Function

Public Function GetPrice(ByVal objBillDetail As BillDetail, ByVal dbl数量 As Double, ByRef dbl成本价 As Double, _
    Optional ByVal lngRow As Long, Optional ByVal blnShowErrMsg As Boolean = True) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品/卫材价格
    '入参:
    '出参:
    '   dbl成本价 成本价
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblPrice As Double, dbl空缺数量 As Double
    
    On Error GoTo ErrHandler
    dbl成本价 = 0
    If objBillDetail Is Nothing Then Exit Function
    
    With objBillDetail
        If .收费类别 = "4" Then
            If mobjService.zlStuffSvr_GetPrice(.收费细目ID, .执行部门ID, dbl数量, .Detail.批次, "", _
                dblPrice, dbl成本价, dbl空缺数量, mlngModule) = False Then
                
                '获取价格失败
                If blnShowErrMsg Then
                    MsgBox IIf(lngRow > 0, "第" & lngRow & "行", "") & "卫生材料""" & .Detail.名称 & """获取价格失败！", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        Else
            If mobjService.zlDrugSvr_GetPrice(.收费细目ID, .执行部门ID, dbl数量, .Detail.批次, "", _
                dblPrice, dbl成本价, dbl空缺数量, mlngModule) = False Then
                
                '获取价格失败
                If blnShowErrMsg Then
                    MsgBox IIf(lngRow > 0, "第" & lngRow & "行", "") & "药品""" & .Detail.名称 & """获取价格失败！", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        End If
        
        If dbl空缺数量 <> 0 And .Detail.变价 Then
            '数量未分解完毕
            If blnShowErrMsg Then
                If InStr(",5,6,7,", .收费类别) > 0 Then
                    MsgBox IIf(lngRow > 0, "第" & lngRow & "行", "") & "时价药品""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                Else
                    MsgBox IIf(lngRow > 0, "第" & lngRow & "行", "") & "时价卫生材料""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        End If
    End With
    GetPrice = dblPrice
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function CheckValidity(ByVal lng材料ID As Long, ByVal lng库房ID As Long, ByVal dbl数量 As Double, _
    Optional ByVal blnAsk As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查卫生材料的灭菌效期是否过期
    '入参:
    '   blnAsk=表示是否询问是否继续,否则为提醒
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    If mobjService.zlStuffSvr_CheckValidity(lng材料ID, lng库房ID, dbl数量, blnAsk, mlngModule) = False Then Exit Function
    CheckValidity = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetStockCheck(ByVal bln卫材 As Boolean) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品/卫材的库存检查方式
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllStockCheck As Collection
    Dim colStock As Collection
    
    On Error GoTo ErrHandler
    If bln卫材 Then
        If mobjService.zlStuffSvr_GetStockCheck(cllStockCheck, mlngModule) = False Then GoTo ReInit:
    Else
        If mobjService.zlDrugSvr_GetStockCheck(cllStockCheck, mlngModule) = False Then GoTo ReInit:
    End If
    
    Err = 0: On Error Resume Next
    cllStockCheck.Add 0, "_0"
    Set GetStockCheck = cllStockCheck
    On Error GoTo 0
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
ReInit:
    Set colStock = New Collection
    colStock.Add 0, "_0" '避免出错
    Set GetStockCheck = colStock
End Function

Public Function CheckItemUnderSet(ByVal lng项目id As Long, ByVal lng库房ID As Long, _
    ByVal dbl库存 As Double, Optional bln卫材 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定药品/材料在指定库房的库存是否低于储备下限
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnBelowLimitLower As Boolean
    
    On Error GoTo ErrHandler
    If bln卫材 Then
        If mobjService.zlStuffSvr_CheckStoreLimit(lng项目id, lng库房ID, dbl库存, blnBelowLimitLower, mlngModule) = False Then Exit Function
    Else
        If mobjService.zlDrugSvr_CheckStoreLimit(lng项目id, lng库房ID, dbl库存, blnBelowLimitLower, mlngModule) = False Then Exit Function
    End If
    CheckItemUnderSet = blnBelowLimitLower
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetStockInfo(ByVal lng项目id As Long, ByVal dbl换算系数 As Double) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品在各个药房，药库的库存信息
    '入参:
    '出参:
    '返回:
    '说明:bln药房/bln药库 至少要有一个设置为真
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str药房性质 As String, strStockInfo As String
    
    On Error GoTo ErrHandler
    Call zlInit缺省部门
    
    If gbln其它药房 And gbln其它药库 Then
        str药房性质 = "中药房,西药房,成药房,中药库,西药库,成药库"
    ElseIf gbln其它药房 Then
        str药房性质 = "中药房,西药房,成药房"
    ElseIf gbln其它药库 Then
        str药房性质 = "中药库,西药库,成药库"
    Else
        Exit Function
    End If
    
    If mobjService.zlDrugSvr_GetStockInfo(lng项目id, str药房性质, dbl换算系数, strStockInfo, mlngModule, _
        gbyt库存显示方式, gstr所属部门ID) = False Then Exit Function
    If strStockInfo = "" Then Exit Function
    
    If zlStr.IsHavePrivs(mstrPrivs, "显示库存") Then
        GetStockInfo = "库存:" & strStockInfo
    Else
        GetStockInfo = "有库存."
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetStock(ByVal lng项目id As Long, ByVal lng库房ID As Long, _
    Optional ByVal dbl药房包装 As Double, Optional ByVal bln卫材 As Boolean, Optional ByVal lng批次 As Long = -1) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定药房指定药品库存(以零售单位)
    '入参:
    '   objBillDetail 单据项目
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl可用库存 As Double
    
    On Error GoTo ErrHandler
    If bln卫材 Then
        If mobjService.zlStuffSvr_GetStock(lng项目id, lng库房ID, lng批次, dbl可用库存, mlngModule) = False Then Exit Function
    Else
        If mobjService.zlDrugSvr_GetStock(lng项目id, lng库房ID, lng批次, dbl可用库存, mlngModule) = False Then Exit Function
    End If
    If Not bln卫材 And gbln药房单位 Then dbl可用库存 = dbl可用库存 / dbl药房包装
    GetStock = dbl可用库存
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Getcargospace(ByVal lng项目id As Long, ByVal lng库房ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定药品在指定药房的库房货位
    '入参:
    '   lng项目ID 药品ID/材料ID
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCargospace As String
    
    On Error GoTo ErrHandler
    If mobjService.zlDrugSvr_Getcargospace(lng项目id, lng库房ID, strCargospace, mlngModule) = False Then Exit Function
    Getcargospace = strCargospace
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetStockByDrugName(ByVal lng药名ID As Long, ByVal lng药房ID As Long, _
    ByVal int显示单位 As Integer, ByRef cllStockData_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据药名,获取指定库存信息
    '入参:
    '   lng药房ID-药房ID,=0时表示不区分药房
    '   int显示单位-0-售价单位;1-门诊单位;2-住院单位
    '出参:
    '   cllStockData_Out-返回数据集，每个成员的key(pharmacy_id,drug_id,stock)
    '返回:成功返回true,否则返回False
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    If mobjService.ZlDrugsvr_GetStockByDrugName(lng药名ID, lng药房ID, int显示单位, cllStockData_Out, 1, mlngModule) = False Then Exit Function
    GetStockByDrugName = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckBill存在非散装草药(objBill As ExpenseBill, ByVal intPage As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中是否存在非散装草药形态
    '入参:
    '   objBill 费用单据对象
    '   intPage 指定的页
    '入参:
    '出参:
    '返回:存在,返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo ErrHandler
    If objBill Is Nothing Then Exit Function
    With objBill.Pages(intPage)
        If .Details.Count = 0 Then Exit Function
        For i = 1 To .Details.Count
            If .Details(i).收费类别 = "7" Then
                If .Details(i).Detail.中药形态 <> 0 Then    '0-散装;1-中药饮片;2-免煎剂
                    CheckBill存在非散装草药 = True: Exit Function
                End If
            End If
        Next
    End With
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get中药形态(objBill As ExpenseBill, ByVal intPage As Integer, _
    Optional ByVal lngRow As Long = -1, Optional blnOnly中成药 As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据中的中药形态
    '入参:
    '   objBill 费用单据对象
    '   intPage 当前第几页
    '   lngRow-当前操作的行
    '   blnOnly中成药-仅判断是否有中成药(对配方时判断有效):原因是中划药在配方中已经存在,就不需要检查
    '出参:
    '返回:录入了中草药的,则返回免煎属性(1-免煎,0-需要煎),否则返回-1 表示还没有录入免煎项目
    '返回:
    '说明:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    
    On Error GoTo ErrHandler
    Get中药形态 = -1
    '如果未指定页,则用当前页
    If objBill Is Nothing Then Exit Function
    strTemp = IIf(blnOnly中成药, ",6,", ",6,7,")
    With objBill.Pages(intPage).Details
        For i = 1 To .Count
            If InStr(1, strTemp, "," & .Item(i).收费类别 & ",") > 0 And .Item(i).收费细目ID <> 0 And i <> lngRow Then
                Get中药形态 = .Item(i).Detail.中药形态
                Exit Function
            End If
        Next
    End With
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowStatusStock(objPanel As Panel, _
    ByVal lng库房ID As Long, ByVal str药品名称 As String, ByVal dbl库存 As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:在状态栏显示药品或卫材的库存
    '入参:
    '出参:
    '返回:
    '说明:31936
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call zlInit缺省部门
    
    If zlStr.IsHavePrivs(mstrPrivs, "显示库存") Then
        If InStr(gstr所属部门ID & ",", "," & lng库房ID & ",") > 0 Or gbyt库存显示方式 <= 0 Then
            objPanel.Text = "[" & str药品名称 & "]可用库存:" & dbl库存
        Else
            objPanel.Text = "[" & str药品名称 & "]可用库存:" & IIf(dbl库存 > 0, "有", "无") & "库存."
        End If
    Else
        objPanel.Text = "[" & str药品名称 & "]" & IIf(dbl库存 > 0, "有", "无") & "库存."
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowStatusCargoSpace(objPanel As Panel, _
    ByVal lng收费细目ID As Long, lng执行库房ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示库房货位
    '入参:
    '出参:
    '返回:
    '说明:27505
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static lngPre收费细目ID As Long
    Static lngPre执行库房ID As Long
    Static strCargo_Space As String  '上次货位
    Dim strTemp As String
    
    On Error GoTo Errhand
    '划价时要显示库房货位
    If Not (lngPre收费细目ID = lng收费细目ID And lng执行库房ID = lngPre执行库房ID) Then
        lngPre收费细目ID = lng收费细目ID: lngPre执行库房ID = lng执行库房ID
        strCargo_Space = Getcargospace(lng收费细目ID, lng执行库房ID)        '重新获取库房货位
    End If
    If strCargo_Space <> "" And InStr(strCargo_Space, "货位:") = 0 Then strCargo_Space = "货位:" & strCargo_Space
    strTemp = Split(objPanel.Text, ",货位:")(0)
    strTemp = Split(strTemp, "货位:")(0)
    If strTemp <> "" And strCargo_Space <> "" Then strTemp = strTemp & ","
    strTemp = strTemp & strCargo_Space
    objPanel.Text = strTemp    '显示出货位
Errhand:
End Sub

Public Function GetDrugTotal(ByVal objBill As ExpenseBill, _
    ByVal lng药品ID As Long, ByVal lng药房ID As Long, _
    Optional ByVal intPage As Integer, Optional ByVal lng批次 As Long = -1) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据中指定药品在同一药房多行的数量合
    '入参:
    '   lng药房ID-0表示分离发药时,不限定药房检查
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, p As Integer, dblCount As Double
    
    On Error GoTo ErrHandler
    If objBill Is Nothing Then Exit Function
    For p = 1 To objBill.Pages.Count
        If intPage = 0 Or p = intPage Then
            For i = 1 To objBill.Pages(p).Details.Count
                If objBill.Pages(p).Details(i).收费细目ID = lng药品ID And _
                    (lng药房ID = 0 Or objBill.Pages(p).Details(i).执行部门ID = lng药房ID) And _
                    (lng批次 <= 0 Or objBill.Pages(p).Details(i).Detail.批次 = lng批次) Then
                    
                    dblCount = dblCount + objBill.Pages(p).Details(i).付数 * objBill.Pages(p).Details(i).数次
                End If
            Next
        End If
    Next
    GetDrugTotal = dblCount
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetOriginalTotal(ByVal objBill As ExpenseBill, _
    ByVal lng药品ID As Long, ByVal lng药房ID As Long, _
    Optional ByVal intPage As Integer) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据中指定药品在同一药房多行的原始数量和
    '入参:
    '   lng药房ID-0表示分离发药时,不限定药房检查
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, p As Integer, dblCount As Double
    
    For p = 1 To objBill.Pages.Count
        If intPage = 0 Or p = intPage Then
            For i = 1 To objBill.Pages(p).Details.Count
                If objBill.Pages(p).Details(i).收费细目ID = lng药品ID Then
                    If (lng药房ID = 0 Or objBill.Pages(p).Details(i).原始执行部门ID = lng药房ID) Then
                        
                        dblCount = dblCount + objBill.Pages(p).Details(i).原始数量
                    End If
                End If
            Next
        End If
    Next
    GetOriginalTotal = dblCount
End Function

Public Function Get处方限量(lng药品ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定药品的处方限量,以零售单位返回
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler
    If mbln处方限量检查 = False Then Exit Function
    
    strSql = "Select Nvl(A.处方限量,0) as 处方限量 From 药品特性 A,药品规格 B Where A.药名ID=B.药名ID And B.药品ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng药品ID)
    If Not rsTmp.EOF Then Get处方限量 = rsTmp!处方限量
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get处方职务(lng药品ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据药品ID获取其处方职务
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    Get处方职务 = "00"
    If mbln处方职务检查 = False Then Exit Function
    
    strSql = "Select Nvl(B.处方职务,'00') as 处方职务 From 药品规格 A,药品特性 B Where A.药名ID=B.药名ID And A.药品ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng药品ID)
    If Not rsTmp.EOF Then Get处方职务 = rsTmp!处方职务
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckDuty(objBill As ExpenseBill, Optional objDetail As Detail, Optional ByVal intCurrPage As Integer, _
    Optional ByVal blnCommon As Boolean = True, Optional ByRef intPage As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定药品行的职务是否与当前医生的职务相匹配
    '入参:
    '   objDetail 正在输入的项目,不传为所有单据所有行
    '   intCurrPage 当前页，objDetail传入时有效
    '   blnCommon 是否正常的判断,否则为医保或公费病人的判断
    '出参:
    '返回:不匹配的行,0为正确,intPage=单据页号
    '说明:职务：1=正高,2=副高,3=中级,4=助理/师级,5=员/士,9=待聘
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, p As Integer, strTmp As String
    Dim int职务A As Integer, int职务B As Integer
    Dim strMsg As String
    
    If mbln处方职务检查 = False Then Exit Function
    
    strTmp = "正高,副高,中级,助理/师级,员/士,,,,待聘"
    intPage = 0
    
    If objDetail Is Nothing Then
        For p = 1 To objBill.Pages.Count
            If objBill.Pages(p).开单人 <> "" Then
                '每张单据开单人不同,当前单据的开单的人职务
                Call GetOperatorInfo(objBill.Pages(p).开单人, , int职务A)
                
                For i = 1 To objBill.Pages(p).Details.Count
                    If InStr(",5,6,7,", objBill.Pages(p).Details(i).收费类别) > 0 Then
                        If objBill.Pages.Count > 1 Then strMsg = "在单据 " & p & "中"
                        If Not blnCommon Then
                            int职务B = Val(Right(objBill.Pages(p).Details(i).Detail.处方职务, 1))
                            If int职务B > 0 Then
                                If int职务A = 0 Then
                                    strMsg = "对医保或公费" & gstrCustomerAppellation & "," & strMsg & _
                                        "第 " & p & " 页 " & i & " 行药品""" & objBill.Pages(p).Details(i).Detail.名称 & _
                                        """要求开单人职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """," & _
                                        "而""" & objBill.Pages(p).开单人 & """未设置职务！"
                                    CheckDuty = 1: intPage = p
                                ElseIf int职务B < int职务A Then
                                    strMsg = "对医保或公费" & gstrCustomerAppellation & "," & strMsg & _
                                        "第 " & p & " 页 " & i & " 行药品""" & objBill.Pages(p).Details(i).Detail.名称 & _
                                        """要求开单人职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上," & _
                                        "而""" & objBill.Pages(p).开单人 & """职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                                    CheckDuty = i: intPage = p: Exit For
                                End If
                            End If
                        Else
                            int职务B = Val(Left(objBill.Pages(p).Details(i).Detail.处方职务, 1))
                            If int职务B > 0 Then
                                If int职务A = 0 Then
                                    strMsg = strMsg & "第 " & p & " 页 " & i & " 行药品""" & objBill.Pages(p).Details(i).Detail.名称 & _
                                        """要求开单人职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """," & _
                                        "而""" & objBill.Pages(p).开单人 & """未设置职务！"
                                    CheckDuty = 1: intPage = p
                                ElseIf int职务B < int职务A Then
                                    strMsg = strMsg & "第 " & p & " 页 " & i & " 行药品""" & objBill.Pages(p).Details(i).Detail.名称 & _
                                        """要求开单人职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上," & _
                                        "而""" & objBill.Pages(p).开单人 & """职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                                    CheckDuty = i: intPage = p: Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next
    ElseIf objBill.Pages(intCurrPage).开单人 <> "" Then
        If InStr(",5,6,7,", objDetail.类别) = 0 Then Exit Function
        Call GetOperatorInfo(objBill.Pages(intCurrPage).开单人, , int职务A)
        
        If Not blnCommon Then
            int职务B = Val(Right(objDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strMsg = "对医保或公费" & gstrCustomerAppellation & ",药品""" & objDetail.名称 & _
                        """要求开单人职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """," & _
                        "而当前开单人未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strMsg = "对医保或公费" & gstrCustomerAppellation & ",药品""" & objDetail.名称 & _
                        """要求开单人职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上," & _
                        "而当前开单人职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        Else
            int职务B = Val(Left(objDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strMsg = "药品""" & objDetail.名称 & """要求开单人职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """," & _
                        "而当前开单人未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strMsg = "药品""" & objDetail.名称 & """要求开单人职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上," & _
                        "而当前开单人职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    
    If CheckDuty > 0 Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Public Function CheckDrugExist(objBill As ExpenseBill, objDetail As Detail, _
    ByVal intPage As Integer, ByVal intRow As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定药品或卫材(跟踪在用)在单据中是否已经存在
    '入参:
    '   objDetail=项目
    '   intPage=要判断的页
    '   intRow=要判断的行
    '出参:
    '返回:
    '说明:时价或分批在同一执行科室禁止重复输入(这里仅提示,保存时禁止)
    '     非时价的分批药品，在不同的单据上有相同的，允许不合并，不提醒
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, p As Integer
    Dim strTmp As String
    
    For p = 1 To objBill.Pages.Count
        For i = 1 To objBill.Pages(p).Details.Count
            If Not (p = intPage And i = intRow) And InStr(",4,5,6,7,", objBill.Pages(p).Details(i).收费类别) > 0 Then
                If objBill.Pages(p).Details(i).Detail.ID = objDetail.ID Then
                    If objBill.Pages.Count > 1 Then strTmp = "第 " & p & " 张单据"
                    If (objBill.Pages(p).Details(i).Detail.分批 Or objBill.Pages(p).Details(i).Detail.变价) _
                        And (objDetail.分批 Or objDetail.变价) Then
                        
                        '非时价的分批药品，在不同的单据上有相同的，允许不合并，不提醒
                        If objDetail.变价 Or (Not objDetail.变价 And objDetail.分批 And intPage = p) Then
                            If objDetail.类别 = "4" Then
                                If MsgBox("卫生材料""" & objDetail.名称 & """在" & strTmp & "第 " & i & " 行已经输入,要继续吗？" & _
                                    vbCrLf & vbCrLf & "注意：该卫生材料为分批或时价材料,重复输入时必须保证它们的发料部门不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    CheckDrugExist = True
                                End If
                            Else
                                If MsgBox("药品""" & objDetail.名称 & """在" & strTmp & "第 " & i & " 行已经输入,要继续吗？" & _
                                    vbCrLf & vbCrLf & "注意：该药品为分批或时价药品,重复输入时必须保证它们的执行药房不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    CheckDrugExist = True
                                End If
                            End If
                            Exit Function
                        End If
                    Else
                        If objDetail.类别 = "4" Then
                            If MsgBox("卫生材料""" & objDetail.名称 & """在" & strTmp & "第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                CheckDrugExist = True
                            End If
                        Else
                            If MsgBox("药品""" & objDetail.名称 & """在" & strTmp & "第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                CheckDrugExist = True
                            End If
                        End If
                        Exit Function
                    End If
                End If
            End If
        Next
    Next
End Function

Public Function CheckLimit(ByVal objBill As ExpenseBill, Optional ByVal intPage As Integer, Optional ByVal intRow As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:费用单据药品处方限量检查,适用于记帐单,收费
    '入参:
    '   intPage,intRow=指定对某张单据某行进行检查,否则为全部检查
    '出参:
    '返回:
    '说明:
    '   1.全部没超过限量，返回真；如有超过药品，则在函数内提示，并返回假
    '   2.记帐表是为每个病人单独检查
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, strSql As String
    Dim tmpDetail As BillDetail, curDetail As BillDetail
    Dim dblTime As Double, strItemIDs As String '已经检查过了的药品
    Dim dbl剂量 As Double, i As Integer, p As Integer
    Dim str药品限量提示 As String
    
    On Error GoTo ErrHandler
    CheckLimit = True
    If mbln处方限量检查 = False Then Exit Function
    
    For p = IIf(intPage = 0, 1, intPage) To IIf(intPage = 0, objBill.Pages.Count, intPage)
        '收集病人
        strItemIDs = ""
        For i = 1 To objBill.Pages(p).Details.Count
            If intRow = 0 Or (intRow > 0 And i = intRow) Then
                With objBill.Pages(p).Details(i)
                    '收集药品ID
                    If InStr(strItemIDs & ",", "," & .收费细目ID & ",") = 0 And InStr(",5,6,7,", .收费类别) > 0 Then
                        strItemIDs = strItemIDs & "," & .收费细目ID
                    End If
                End With
            End If
        Next
        If strItemIDs <> "" Then
            strItemIDs = Mid(strItemIDs, 2)
            strSql = "Select A.药品ID,A.剂量系数,B.计算单位 as 剂量单位" & _
                " From 药品规格 A,诊疗项目目录 B" & _
                " Where A.药名ID=B.ID And A.药品ID IN (" & strItemIDs & ")"
            Call zlDatabase.OpenRecordset(rsTmp, strSql, "mdlOutExse")
            strItemIDs = ""
            For i = 1 To objBill.Pages(p).Details.Count
                If intRow = 0 Or (intRow > 0 And i = intRow) Then
                    Set tmpDetail = objBill.Pages(p).Details(i)
                    If InStr(",5,6,7,", tmpDetail.收费类别) > 0 And tmpDetail.Detail.处方限量 > 0 Then
                        If InStr(strItemIDs, "," & tmpDetail.收费细目ID) = 0 Then
                            dblTime = 0
                            For Each curDetail In objBill.Pages(p).Details
                                If InStr(",5,6,7,", curDetail.收费类别) > 0 And tmpDetail.收费细目ID = curDetail.收费细目ID Then
                                    dblTime = dblTime + curDetail.付数 * curDetail.数次
                                End If
                            Next
                            rsTmp.Filter = "药品ID=" & tmpDetail.收费细目ID
                            If Not rsTmp.EOF Then
                                If gbln药房单位 Then
                                    dbl剂量 = dblTime * tmpDetail.Detail.药房包装 * rsTmp!剂量系数
                                Else
                                    dbl剂量 = dblTime * rsTmp!剂量系数
                                End If
                                If dbl剂量 > tmpDetail.Detail.处方限量 Then
                                    str药品限量提示 = IIf(objBill.Pages.Count = 1, "", "单据" & p & "中") & "药品 """ & tmpDetail.Detail.名称 & """ 的总剂量 " & _
                                        FormatEx(dbl剂量, 5) & rsTmp!剂量单位 & "(" & FormatEx(dblTime, 5) & IIf(gbln药房单位, tmpDetail.Detail.药房单位, tmpDetail.Detail.计算单位) & ") 超过处方限量 " & _
                                        FormatEx(tmpDetail.Detail.处方限量, 5) & rsTmp!剂量单位 & " ！" & vbCrLf & vbCrLf & "是否允许?允许并且退出窗口之前不再进行处方限量检查请选择[取消]"
                                
                                        str药品限量提示 = MsgBox(str药品限量提示, vbYesNoCancel + vbDefaultButton3 + vbInformation, gstrSysName)
                                        If str药品限量提示 = vbNo Then
                                            CheckLimit = False: Exit Function
                                        End If
                                        If str药品限量提示 = vbCancel Then
                                           mbln已检查处方限量 = True
                                        End If
                                End If
                            End If
                            strItemIDs = strItemIDs & "," & tmpDetail.收费细目ID
                        End If
                    End If
                End If
            Next
        End If
    Next
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckStock(objDetail As Detail, ByVal lng库房ID As Long, ByVal dbl数量 As Double, _
    Optional ByVal bln卫材 As Boolean, Optional ByRef strBeginStr As String, _
    Optional ByVal blnSaveBefore As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:库存检查
    '入参:
    '   dbl数量 当前检查数量
    '   bln卫材 是否卫生材料
    '   strBeginStr 提示文本前面添加显示文本
    '   blnSaveBefore 是否保存前调用检查，此时如果已提示则不再提示
    '出参:
    '返回:检查通过返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim colStock As Collection, strStock As String

    On Error GoTo ErrHandler
    If dbl数量 <= objDetail.库存 Then CheckStock = True: Exit Function
    
    strStock = FormatEx(objDetail.库存, 5) & IIf(gbln药房单位, objDetail.药房单位, objDetail.计算单位)
    
    '药品库存检查:动态药房,分批或时价药品也要检查了
    If objDetail.分批 Or objDetail.变价 Then '分批或时价药品库存不足禁止输入
        MsgBox strBeginStr & "[" & objDetail.名称 & "]为分批或时价" & IIf(bln卫材, "卫生材料", "药品") & _
            "，当前可用库存" & strStock & "不足输入数量" & FormatEx(dbl数量, 5) & "！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "不检查库存") Then CheckStock = True: Exit Function
    
    Set colStock = IIf(bln卫材, mcolStuffStockCheck, mcolDrugStockCheck)
    
    Select Case colStock("_" & lng库房ID)
    Case StockCheckType.NOCheck
        CheckStock = True: Exit Function
    Case StockCheckType.CheckForbid
        MsgBox strBeginStr & "[" & objDetail.名称 & "]的当前可用库存" & strStock & "不足输入数量" & FormatEx(dbl数量, 5) & _
            "，请修改或检查是否有多行输入。！", vbInformation, gstrSysName
        Exit Function
    Case StockCheckType.CheckRemind
        '保存时如果已检查则不再检查
        If Not blnSaveBefore Or blnSaveBefore And objDetail.库存检查 = CheckRemindSelect.NORemind Then
            If MsgBox(strBeginStr & "[" & objDetail.名称 & "]的当前可用库存" & strStock & "不足输入数量" & FormatEx(dbl数量, 5) & _
                "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                objDetail.库存检查 = CheckRemindSelect.Ban
                Exit Function
            Else
                objDetail.库存检查 = CheckRemindSelect.Allow
            End If
        End If
    End Select
    CheckStock = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckDisable(objBill As ExpenseBill) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中的药品的禁忌情况
    '入参:
    '出参:
    '返回:药品互相禁忌提示信息
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strInfo As String
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim strGroup As String, strIDs As String
    Dim blnStop As Boolean, strMsg As String
    
    On Error GoTo ErrHandler
    For p = 1 To objBill.Pages.Count
        strInfo = "": strIDs = ""
        For i = 1 To objBill.Pages(p).Details.Count
            If InStr(",5,6,7,", objBill.Pages(p).Details(i).收费类别) > 0 Then
                strIDs = strIDs & "," & objBill.Pages(p).Details(i).收费细目ID
            End If
        Next
        strIDs = Mid(strIDs, 2)
        If Not (strIDs = "" Or UBound(Split(strIDs, ",")) < 1) Then
            strSql = _
                " Select /*+Cardinality(j,10)*/ A.组编号,Count(Distinct A.项目ID) as 禁忌数" & _
                " From 诊疗互斥项目 A,药品规格 B,Table(f_num2list([1])) J" & _
                " Where A.项目ID=B.药名ID And B.药品ID  = j.Column_Value" & _
                " Having Count(Distinct A.项目ID)>1  " & _
                "  Group by A.组编号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlOutExse", strIDs)
            
            If Not rsTmp.EOF Then
                strGroup = ""
                For i = 1 To rsTmp.RecordCount
                    strGroup = strGroup & "," & rsTmp!组编号
                    rsTmp.MoveNext
                Next
                strGroup = Mid(strGroup, 2)
                
                For i = 0 To UBound(Split(strGroup, ","))
                    strSql = _
                        "Select /*+Cardinality(j,10)*/ Distinct C.类型,C.组编号,D.编码,D.名称,D.规格" & _
                        " From 药品规格 A,诊疗项目目录 B,诊疗互斥项目 C,收费项目目录 D,Table(f_num2list([1])) J" & _
                        " Where A.药名ID=B.ID And B.ID=C.项目ID And A.药品ID=D.ID" & _
                        "           And C.组编号=[1]" & _
                        "           And A.药品ID  = j.Column_Value" & _
                        " Order by C.类型,C.组编号,D.编码"
                        
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlOutExse", Val(Split(strGroup, ",")(i)), strIDs)
                    If Not rsTmp.EOF Then
                        rsTmp.Filter = "类型=1"
                        If rsTmp.RecordCount > 1 Then
                            k = k + 1
                            strInfo = strInfo & vbCrLf & "第 " & k & " 组(互相慎用)：" & vbCrLf
                            For j = 1 To rsTmp.RecordCount
                                strInfo = strInfo & "[" & rsTmp!编码 & "]" & rsTmp!名称 & IIf(IsNull(rsTmp!规格), "", "(" & rsTmp!规格 & ")") & "                 " & vbCrLf
                                rsTmp.MoveNext
                            Next
                        End If
                        rsTmp.Filter = "类型=2"
                        If rsTmp.RecordCount > 1 Then
                            blnStop = True
                            k = k + 1
                            strInfo = strInfo & vbCrLf & "第 " & k & " 组(互相禁用)：" & vbCrLf
                            For j = 1 To rsTmp.RecordCount
                                strInfo = strInfo & "[" & rsTmp!编码 & "]" & rsTmp!名称 & IIf(IsNull(rsTmp!规格), "", "(" & rsTmp!规格 & ")") & "                 " & vbCrLf
                                rsTmp.MoveNext
                            Next
                        End If
                        rsTmp.Filter = 0
                    End If
                Next
                If strInfo <> "" Then
                    If objBill.Pages.Count = 1 Then
                        strMsg = strMsg & vbCrLf & "单据中下列药品互相禁用或慎用：" & vbCrLf & strInfo
                    Else
                        strMsg = strMsg & vbCrLf & "单据" & p & "中下列药品互相禁用或慎用：" & vbCrLf & strInfo
                    End If
                End If
            End If
        End If
    Next
    If strMsg <> "" Then
        If blnStop Then
            strMsg = strMsg & vbCrLf & "请修改禁用药品后再继续！"
        Else
            strMsg = strMsg & vbCrLf & "要继续吗？"
        End If
    End If
    strMsg = Mid(strMsg, 3)
    
    If strMsg <> "" Then
        If strMsg Like "*(互相禁用)*" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        Else
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    CheckDisable = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckExistRepeatRow(objBill As ExpenseBill, ByRef intPage As Integer, ByRef blnMerge As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查分批或时价药品同一药房是否有重复输入
    '入参:
    '出参:
    '   blnMerge 是否存在需要合并行
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    Dim p As Integer, i As Integer, j As Integer, k As Integer
    
    On Error GoTo ErrHandler
    intPage = 1: blnMerge = False
    If objBill Is Nothing Then Exit Function
    
    blnMerge = False
    For p = 1 To objBill.Pages.Count
        For i = 1 To objBill.Pages(p).Details.Count
            With objBill.Pages(p).Details(i)
                If (.Detail.分批 Or .Detail.变价) And (InStr(",5,6,7,", .收费类别) > 0 Or .收费类别 = "4" And .Detail.跟踪在用) Then
                    For k = 1 To objBill.Pages.Count
                        For j = 1 To objBill.Pages(k).Details.Count
                            If Not (p = k And i = j) And .收费细目ID = objBill.Pages(k).Details(j).收费细目ID And .执行部门ID = objBill.Pages(k).Details(j).执行部门ID Then
                                '多张单据的情况
                                If objBill.Pages.Count > 1 Then
                                    '非时价的分批药品，在不同的单据上有相同的，允许不合并，不提醒
                                    If .Detail.变价 Or (Not .Detail.变价 And .Detail.分批 And p = k) Then
                                        If .收费类别 = "4" Then
                                            If Not blnMerge Then
                                                If MsgBox("第 " & p & " 张单据第 " & i & " 行,及第 " & k & " 张单据第 " & j & " 行的" & _
                                                    vbCrLf & "分批或时价卫生材料""" & .Detail.名称 & """在同一个发料部门被重复输入。" & _
                                                    vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？", _
                                                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                    blnMerge = True     '不应退出循环，因为还要检查是否有不同付数的中草药,如果有的话，不能自动合并
                                                Else
                                                    intPage = k: Exit Function
                                                End If
                                            End If
                                        Else
                                            '两个不同单据中的中药付数不同时，应是不同的配方，无法自动合并
                                            If .收费类别 = "7" And .付数 <> objBill.Pages(k).Details(j).付数 Then
                                                MsgBox "第 " & p & " 张单据第 " & i & " 行,及第 " & k & " 张单据第 " & j & " 行的" & _
                                                    vbCrLf & "分批或时价中草药""" & .Detail.名称 & """(不同付数)在同一个药房被重复输入。", vbInformation, gstrSysName
                                                intPage = k: Exit Function
                                            ElseIf Not blnMerge Then
                                                If MsgBox("第 " & p & " 张单据第 " & i & " 行,及第 " & k & " 张单据第 " & j & " 行的" & _
                                                    vbCrLf & "分批或时价药品""" & .Detail.名称 & """在同一个药房被重复输入。" & _
                                                    vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？", _
                                                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                    blnMerge = True
                                                Else
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                    End If
                                '单张单据的情况
                                ElseIf Not blnMerge Then
                                    If .收费类别 = "4" Then
                                        strInfo = "第 " & j & " 行的分批或时价卫生材料""" & .Detail.名称 & """在同一个发料部门被重复输入。" & _
                                                    vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？"
                                    Else
                                        strInfo = "第 " & j & " 行的分批或时价药品""" & .Detail.名称 & """在同一个药房被重复输入。" & _
                                                    vbCrLf & vbCrLf & "要自动合并单据中所有重复输入的分批或时价项目吗？"
                                    End If
                                    
                                    If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                        blnMerge = True     '可以退出循环
                                    Else
                                        intPage = 1: Exit Function
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If
            End With
        Next
    Next
    CheckExistRepeatRow = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckServiceDept(objBill As ExpenseBill) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查药品的发药药房对应的服务科室
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, p As Integer
    Dim strTmp As String, lng收费细目ID As Long
    
    On Error GoTo ErrHandler
    '收集药品的发药药房对应的服务科室
    strTmp = ""
    For p = 1 To objBill.Pages.Count
        For i = 1 To objBill.Pages(p).Details.Count
            If InStr(",5,6,7,", objBill.Pages(p).Details(i).收费类别) > 0 Then
                strTmp = strTmp & "," & objBill.Pages(p).Details(i).收费细目ID
            End If
        Next
    Next
    If strTmp = "" Then CheckServiceDept = True: Exit Function
    
    Set rsTmp = GetServiceDept(strTmp)
    If rsTmp Is Nothing Then CheckServiceDept = True: Exit Function
    
    For p = 1 To objBill.Pages.Count
        strTmp = ""
        For i = 1 To objBill.Pages(p).Details.Count
            If InStr(",5,6,7,", objBill.Pages(p).Details(i).收费类别) > 0 Then
                lng收费细目ID = objBill.Pages(p).Details(i).收费细目ID
                '先检查是否是允许的存储库房
                rsTmp.Filter = "收费细目ID=" & lng收费细目ID & " And 执行科室id=" & objBill.Pages(p).Details(i).执行部门ID
                If rsTmp.RecordCount = 0 Then
                    strTmp = strTmp & "," & i
                Else
                    '再检查是否是允许的服务科室(没有设置服务科室的,开单科室ID为零)
                    rsTmp.Filter = "(" & rsTmp.Filter & " And 开单科室ID=" & _
                        IIf(objBill.科室ID = 0, objBill.Pages(p).开单部门ID, objBill.科室ID) & ") Or (" & rsTmp.Filter & " And 开单科室ID=0)"
                    If rsTmp.RecordCount = 0 Then
                        strTmp = strTmp & "," & i
                    End If
                End If
            End If
        Next
        If strTmp <> "" Then
            strTmp = Mid(strTmp, 2)
            MsgBox "请检查,第" & p & "张单据,第" & strTmp & "行药品是否违反以下规则:" & vbCrLf & vbCrLf & _
                "A.选择的执行科室不是药品的存储库房" & vbCrLf & _
                "B.病人科室[" & GET部门名称(IIf(objBill.科室ID = 0, objBill.Pages(p).开单部门ID, objBill.科室ID)) & "]不属于药品在此存储库房的服务科室.", _
                vbInformation, gstrSysName
            Exit Function
        End If
    Next
    CheckServiceDept = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPriceAdjustBySell(objBill As ExpenseBill, ByRef intPage As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, p As Integer
    
    On Error GoTo ErrHandler
    If mobjPublicDrug Is Nothing Then CheckPriceAdjustBySell = True: Exit Function
    
    'Private Function zlCheckPriceAdjustBySell(ByVal lng药品id As Long, ByVal lng药房id As Long) As Boolean
    '零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
    '定价药品：售价是固定的，比较所有药房的成本价，如果存在不一致的就不能销售出库
    '时价药品：比较药房库存记录的零售价和成本价，如果存在不一致的就不能销售出库
    '销售出库时只判断药房
    '返回：True-正常进行销售出库；false-不能进行销售出库
    For p = 1 To objBill.Pages.Count
        For i = 1 To objBill.Pages(p).Details.Count
            With objBill.Pages(p).Details(i)
                If InStr(",5,6,7,", .收费类别) > 0 Then
                    If mobjPublicDrug.zlCheckPriceAdjustBySell(.收费细目ID, .执行部门ID) = False Then
                        intPage = p: Exit Function
                    End If
                End If
            End With
        Next
    Next
    CheckPriceAdjustBySell = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateDrugPacker() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建自助发药机(自动化药房)
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String, strMessage As String
    Dim blnSuccess As Boolean

    If Not mobjDrugMachine Is Nothing Or Not mobjDrugPacker Is Nothing Then CreateDrugPacker = True: Exit Function

    On Error Resume Next
    If Val(zlDatabase.GetPara("启用药品自动化设备接口", glngSys, Val("9010-药品自动化设备接口"))) = 1 Then
        '优先新接口
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err <> 0 Then Set mobjDrugMachine = Nothing
        
        On Error GoTo 0
        If Not mobjDrugMachine Is Nothing Then
            '权限检查
            strPrivs = GetPrivFunc(glngSys, Val("9010-药品自动化设备接口"))
            If zlStr.IsHavePrivs(strPrivs, "基本") = False Then Set mobjDrugMachine = Nothing: Exit Function
            blnSuccess = mobjDrugMachine.Init(1, objComLib, strMessage)
            If blnSuccess = False Then Set mobjDrugMachine = Nothing: Exit Function
            CreateDrugPacker = True: Exit Function
        End If
    End If
    
    '旧部件
    On Error Resume Next
    Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
    If Err <> 0 Then Set mobjDrugMachine = Nothing
    If mobjDrugPacker Is Nothing Then Exit Function
    
    On Error GoTo 0
    blnSuccess = mobjDrugPacker.DYEY_MZ_IniSoap()
    If blnSuccess = False Then Set mobjDrugPacker = Nothing: Exit Function
    
    CreateDrugPacker = True
End Function

Public Sub DrugMachine_Charge(ByVal bytBillType As Byte, ByVal strNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费或记帐后发药机数据上传
    '入参:
    '   bytBillType 单据类型：1-收费单,2-记帐单
    '   strNos 费用单据，格式：A001,A002
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String, strReturn As String
    
    On Error GoTo ErrHandler
    If CreateDrugPacker() = False Then Exit Sub
    
    Select Case bytBillType
    Case 0: '收费单
        If Not mobjDrugMachine Is Nothing Then
            '门诊格式：1|单据1,处方号1;单据2,处方号2
            strData = "1|" & "8," & Replace(strNos, ",", ";8,")
            Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
            Exit Sub
        End If
        
        If mobjDrugPacker Is Nothing Then Exit Sub
        '格式：单据1,处方号1|单据2,处方号2
        strData = "8," & Replace(strNos, ",", "|8,")
        Call mobjDrugPacker.DYEY_MZ_TransRecipeDetail(1, UserInfo.编号, UserInfo.姓名, 0, strData, strReturn)
    Case 2: '记帐单
        If mobjDrugMachine Is Nothing Then Exit Sub
        
        '门诊格式：1|单据1,处方号1;单据2,处方号2
        strData = "1|" & "9," & Replace(strNos, ",", ";9,")
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DrugMachine_Del(ByVal bytBillType As Byte, ByVal varValue As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退费或销账后发药机数据上传
    '入参:
    '   bytBillType 单据类型：1-收费单,2-记帐单
    '   varValue bytBillType=0,结算序号；bytBillType=2,单据号
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsData As ADODB.Recordset
    Dim strReturn As String, strData As String '门诊处方退药格式：费用ID1,退药数量1;费用ID2,退药数量2;...
    Dim lng结算序号 As Long, strNo As String
    
    On Error GoTo ErrHandler
    If CreateDrugPacker() = False Then Exit Sub
    
    Select Case bytBillType
    Case 0: '收费单
        lng结算序号 = varValue
        If Not mobjDrugMachine Is Nothing Then
            '本次退的减去重收的就是实际退的
            strSql = "Select Max(Decode(a.记录状态, 2, a.Id, 0)) As 费用id, -1 * Nvl(Sum(a.付数 * a.数次), 0) As 退药数量" & _
                    " From 门诊费用记录 A,(Select Distinct 结帐ID From 病人预交记录 Where 结算序号 = [1]) B" & _
                    " Where a.结帐id = b.结帐ID And Mod(a.记录性质, 10) = 1 And a.收费类别 In('5','6','7')" & _
                    " Group By NO, Nvl(价格父号, 序号)" & _
                    " Having Nvl(Sum(a.付数 * a.数次), 0) <> 0"
            Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询退费数据", lng结算序号)
            If rsData.RecordCount = 0 Then Exit Sub
            
            Do While Not rsData.EOF
                strData = strData & ";" & Nvl(rsData!费用id) & "," & Nvl(rsData!退药数量)
                rsData.MoveNext
            Loop
            strData = Mid(strData, 2)
            Call mobjDrugMachine.Operation(gstrDBUser, Val("24-处方退药(完整/部分)"), strData, strReturn)
            Exit Sub
        End If
        
        If mobjDrugPacker Is Nothing Then Exit Sub
        strSql = "Select a.No, a.执行部门id" & _
                " From 门诊费用记录 A, 病人预交记录 B" & _
                " Where a.结帐id = b.结帐id And a.记录状态=2  And a.收费类别 In('5','6','7') And b.结算序号 = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询退费数据", lng结算序号)
        If rsData.RecordCount = 0 Then Exit Sub
        
        Do While Not rsData.EOF
            If InStr(strData & "|", "|" & Nvl(rsData!NO) & "," & Nvl(rsData!执行部门ID) & "|") = 0 Then
                strData = strData & "|" & Nvl(rsData!NO) & "," & Nvl(rsData!执行部门ID)
            End If
            rsData.MoveNext
        Loop
        strData = Mid(strData, 2)
        Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.编号, UserInfo.姓名, strData, strReturn)
        
    Case 2: '记帐单
        If mobjDrugMachine Is Nothing Then Exit Sub
        
        strNo = varValue
        strSql = "Select Id As 费用id, -1 * Nvl(付数, 1) * 数次 As 退药数量" & _
                " From 门诊费用记录" & _
                " Where 记录性质 = 2 And 记录状态 = 2 And NO = [1] And 收费类别 In('5','6','7')" & _
                "       And 登记时间 + 0 = (Select Max(登记时间)" & _
                "                       From 门诊费用记录" & _
                "                       Where 记录性质 = 2 And 记录状态 = 2 And NO = [1])"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询销账数据", strNo)
        If rsData.RecordCount = 0 Then Exit Sub
        
        Do While Not rsData.EOF
            strData = strData & ";" & Nvl(rsData!费用id) & "," & Nvl(rsData!退药数量)
            rsData.MoveNext
        Loop
        strData = Mid(strData, 2)
        Call mobjDrugMachine.Operation(gstrDBUser, Val("24-处方退药(完整/部分)"), strData, strReturn)
        
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function Read发药窗口(ByVal lng药房ID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据药房ID获取发药窗口
    '入参:
    '出参:
    '返回:发药窗口记录集：药房ID,发药窗口,是否专家
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsWindows As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mobjService.zlDrugSvr_GetPharmacyWindows(lng药房ID, rsWindows, mlngModule) = False Then Exit Function
    Set Read发药窗口 = rsWindows
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetBillNoWindowDrug(ByVal strNo As String, ByVal bytFlag As Byte) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据中未分配发药窗口的药品的药房ID
    '入参:
    '   strNo 单据号
    '   bytFlag 记录性质
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
     
    On Error GoTo ErrHandler
    strSql = "Select 执行部门id" & _
            " From 门诊费用记录" & _
            " Where 收费类别 In ('5', '6', '7') And NO = [1] And 记录性质 = [2]" & _
            "       And 执行部门id Is Not Null And 发药窗口 Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlOutExse", strNo, bytFlag)
    If rsTmp.EOF Then Exit Function
    GetBillNoWindowDrug = rsTmp!执行部门ID
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckExistWindow(ByVal strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发药窗口检查
    '入参:
    '   strNo 单据号
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng药房ID As Long, rsWindows As ADODB.Recordset
    
    On Error GoTo ErrHandler
    lng药房ID = GetBillNoWindowDrug(strNo, 1)
    If lng药房ID = 0 Then CheckExistWindow = True: Exit Function
    
    Set rsWindows = Read发药窗口(lng药房ID)
    '因为专家窗仅指定,不动态分配,所以排开
    rsWindows.Filter = "药房ID=" & lng药房ID & " And 是否专家=0"
    If Not rsWindows.EOF Then
        MsgBox "无法分配" & GET部门名称(lng药房ID) & "的发药窗口，请确定是否正常安排窗口上班。", vbInformation, gstrSysName
        Exit Function
    End If
    CheckExistWindow = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDrugWindowFromBill(objBill As ExpenseBill, ByVal intPage As Integer, ByVal lng药房ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从已分配发药窗口的单据中取发药窗口
    '入参:
    '   intPage=搜录到的单据编号
    '出参:
    '返回:
    '说明:
    '   主要用于多单据收费时，不同类别的药品可能动态分配到同一药房，这样他们的窗口也应相同，但强行指定的除外
    '   如果是划价单,则以第一药品行的窗口为准,否则以已输入相同药品的窗口为准
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim p As Integer, i As Integer
    Dim varData As Variant, varTemp As Variant
    
    For p = 1 To intPage
        If objBill.Pages(p).NO <> "" Then
            If objBill.Pages(p).发药窗口 <> "" Then
                '问题:47489
                '取划价单的第一药品行的药房进行比较
                '执行部门ID|发药窗口;...
                varData = Split(objBill.Pages(p).发药窗口, ";")
                For i = 0 To UBound(varData)
                    varTemp = Split(varData(i) & "|", "|")
                    If varTemp(0) = lng药房ID Then
                        GetDrugWindowFromBill = varTemp(1)
                        Exit Function
                    End If
                Next
            End If
        Else
            For i = 1 To objBill.Pages(p).Details.Count
                If objBill.Pages(p).Details(i).执行部门ID = lng药房ID _
                    And InStr(",5,6,7,", objBill.Pages(p).Details(i).收费类别) > 0 _
                    And objBill.Pages(p).Details(i).发药窗口 <> "" Then
                     
                    GetDrugWindowFromBill = objBill.Pages(p).Details(i).发药窗口
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Private Function GetDefaultWindow(ByVal str类别 As String, ByVal lng药房ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省的药房窗口设置
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long, arrTmp As Variant, arrWin As Variant
    
    Select Case str类别
        Case "5"
            If InStr(gstr西窗, ":") > 0 Then '旧数据没有存药房ID
                strTmp = gstr西窗
            ElseIf glng西药房 > 0 And gstr西窗 <> "" Then
                strTmp = glng西药房 & ":" & gstr西窗
            End If
        Case "6"
            If InStr(gstr中窗, ":") > 0 Then
                strTmp = gstr中窗
            ElseIf glng中药房 > 0 And gstr中窗 <> "" Then
                strTmp = glng中药房 & ":" & gstr中窗
            End If
        Case "7"
            If InStr(gstr中窗, ":") > 0 Then
                strTmp = gstr成窗
            ElseIf glng成药房 > 0 And gstr成窗 <> "" Then
                strTmp = glng成药房 & ":" & gstr成窗
            End If
    End Select
    
    If strTmp <> "" Then
        arrTmp = Split(strTmp, ",")
        strTmp = ""
        For i = 0 To UBound(arrTmp)
            arrWin = Split(arrTmp(i), ":")
            Select Case str类别
                Case "5"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
                Case "6"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
                Case "7"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
            End Select
        Next
    End If
    GetDefaultWindow = strTmp
End Function

Public Function Get发药窗口(ByVal lng病人ID As Long, ByVal str类别 As String, ByVal lng药房ID As Long, _
    Optional ByVal blnFirst As Boolean, Optional ByVal bytBillType As Byte = 1) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:分配发药窗口
    '入参:
    '   bytBillType 单据类型：1-收费单,2-记帐单
    '出参:
    '返回:分配成功，返回true,否则返回False
    '说明:
    '   根据药房ID来确定,相同的药房ID分配相同的发药窗口
    '   修改时保持原有发药窗口，不会进来
    '规则说明：
    '    1.前面单据和前面行已分配发药窗口，则以前面的为准（保证相同药房有同一个窗口）
    '    2.判断指定病人在指定药房的未发药品记录中是否存在正在上班的发药窗口
    '      a.发药窗口存在，返回填制日期最近的发药窗口
    '      b.发药窗口不存在:
    '        i:如果存在缺省的发药窗口，且正在上班，则返回缺省的发药窗口，如果窗口未上班则返回null
    '        ii:如果不存在缺省的发药窗口，则根据动态分配规则（0-闲忙;1-平均）获取非专家的发药窗口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str发药窗口 As String
    Dim rsData As ADODB.Recordset, str缺省 As String
    
    On Error GoTo ErrHandler
    If blnFirst Then Set mcllWindows = New Collection
     
    If InStr(",5,6,7,", str类别) = 0 Then Exit Function
        
    '分配窗口时，如果发现药房与上张单据不同，则清除缺省窗口,以免药房不同分配到相同窗口
    If str类别 = "5" Then
        If lng药房ID <> mlng西药房 And mlng西药房 <> 0 Then mstr西窗 = ""
        mlng西药房 = lng药房ID '记录该病人使用的药房(划价已定)
    ElseIf str类别 = "6" Then
        If lng药房ID <> mlng成药房 And mlng成药房 <> 0 Then mstr成窗 = ""
        mlng成药房 = lng药房ID
    ElseIf str类别 = "7" Then
        If lng药房ID <> mlng中药房 And mlng中药房 <> 0 Then mstr中窗 = ""
        mlng中药房 = lng药房ID
    End If
    
    '2.从已分配中取:检查是否存在该窗口，保证相同药房有同一个窗口
    For i = 1 To mcllWindows.Count
        If mcllWindows(i)(0) = lng药房ID Then
            Get发药窗口 = mcllWindows(i)(1): Exit Function
        End If
    Next
    
    '3.指定时固定分配(指定是指没有对应药房上班时指定)
    Select Case str类别
    Case "5"
        If mstr西窗 <> "" Then
            str缺省 = mstr西窗
        ElseIf glng西药房 > 0 Then
            str缺省 = GetDefaultWindow(str类别, lng药房ID)
        End If
    Case "6"
        If mstr成窗 <> "" Then
            str缺省 = mstr成窗
        ElseIf glng成药房 > 0 Then
            str缺省 = GetDefaultWindow(str类别, lng药房ID)
        End If
    Case "7"
        If mstr中窗 <> "" Then
            str缺省 = mstr中窗
        ElseIf glng中药房 > 0 Then
            str缺省 = GetDefaultWindow(str类别, lng药房ID)
        End If
    End Select
    
    '4.由药品系统那边分配
    If mobjService.zlDrugSvr_GetSendWindows(1, lng病人ID, IIf(bytBillType = 1, 0, gTy_System_Para.Sy_Reg.bytNODaysGeneral), _
         lng药房ID & "," & str缺省, rsData, mlngModule) = False Then Exit Function
    If Not rsData.EOF Then str发药窗口 = Nvl(rsData!发药窗口)
    
    If str发药窗口 <> "" Then
        Select Case str类别
            Case "5"
                mstr西窗 = str发药窗口
            Case "6"
                mstr成窗 = str发药窗口
            Case "7"
                mstr中窗 = str发药窗口
        End Select
        
        mcllWindows.Add Array(lng药房ID, str发药窗口), "K" & lng药房ID
    End If
    Get发药窗口 = str发药窗口
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Class_Terminate()
    '释放对象
    Set mcolDrugStockCheck = Nothing
    Set mcolStuffStockCheck = Nothing
    
    Set mobjPublicDrug = Nothing
    Set mobjService = Nothing
    
    Set mobjDrugPacker = Nothing
    Set mobjDrugMachine = Nothing
    
    Call RestWindowVariable
End Sub

Public Function AutoSplitSpeci(ByVal lng药名ID As Long, ByVal byt中药形态 As Byte, _
    ByVal int付数 As Integer, ByVal dbl数量 As Double, ByVal lng药房ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:针对中草药根据药名（品种）来自动分配药品(自动分解多个规格)
    '入参:
    '   byt中药形态 0-散装;1-中药饮片;2-免煎剂
    '   byt场合 1-门诊 ，2-住院
    '出参:
    '返回:
    '说明:
    '   返回格式：药品id,数量;药品id,数量;...(散装只选择一个规格)
    '            不能完全分配时返回:剂量为6和10的情况下,17克的分配=23755,6;23756,10|1
    '            不能分配时返回空,例如:剂量为6和10的情况下,3克的分配
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String
    
    On Error GoTo ErrHandler
    If mobjService.zlDrugSvr_AutoSplitSpeci(lng药名ID, byt中药形态, _
        int付数, dbl数量, lng药房ID, strData, 1, mlngModule) = False Then Exit Function
    AutoSplitSpeci = strData
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckIsputDrug(ByVal strNos As String, Optional ByVal bytBillType As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查药品是否已进行配药
    '入参:
    '   strNOs 处方单据，多个用英文逗号分隔
    '   bytBillType 单据类型:1-收费处方；2-记帐单处方；3-记帐表处方
    '出参:
    '   blnIsPutDrug 是否已进行配药
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnIsPutDrug As Boolean
    
    On Error GoTo ErrHandler
    If mbyt药品摆药退费方式 = 0 Then CheckIsputDrug = True: Exit Function
    
    If mobjService.zlDrugSvr_Checkisputdrug(bytBillType, strNos, blnIsPutDrug, mlngModule) = False Then Exit Function
    If blnIsPutDrug = False Then CheckIsputDrug = True: Exit Function
    
    If mbyt药品摆药退费方式 = 1 Then
        '禁止退费
        MsgBox "在退费单据中已经存在摆药的单据，不允许对已经摆药的单据进行退费！", vbInformation, gstrSysName
        Exit Function
    End If
    '提示
    If MsgBox("在退费单据中已经存在摆药的单据，是否对已经摆药的单据进行退费？", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    CheckIsputDrug = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SaveDrugID(objBill As ExpenseBill, intPage As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存当前指定单据序号以前中最后一个含有药品的单据的第一行药品的部门ID
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    
    '记录该病人本次收费分配的各药房(单据内容为输入时)
    For i = 1 To intPage
        If objBill.Pages(i).NO = "" Then
            j = GetFirstRow(objBill, i)
            If j > 0 Then
                Select Case objBill.Pages(i).Details(j).收费类别
                    Case "5"
                        mlng西药房 = objBill.Pages(i).Details(j).执行部门ID
                    Case "6"
                        mlng成药房 = objBill.Pages(i).Details(j).执行部门ID
                    Case "7"
                        mlng中药房 = objBill.Pages(i).Details(j).执行部门ID
                End Select
            End If
        Else
            Call BillDrugDept(objBill.Pages(i).NO)
        End If
    Next
End Sub

Private Function GetFirstRow(objBill As ExpenseBill, intPage As Long, Optional ByVal strClass As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前单据中第一个为药品的行号
    '入参:
    '   strClass=是否只取指定类别药品行
    '出参:
    '返回:0=没有药品收费行
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    For i = 1 To objBill.Pages(intPage).Details.Count
        If strClass = "" Then
            If InStr(",5,6,7,", objBill.Pages(intPage).Details(i).收费类别) > 0 Then
                GetFirstRow = i: Exit Function
            End If
        Else
            If objBill.Pages(intPage).Details(i).收费类别 = strClass Then
                GetFirstRow = i: Exit Function
            End If
        End If
    Next
End Function

Private Function BillDrugDept(ByVal strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取一张收费单据的药房ID
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select 收费类别,执行部门ID " & _
            " From 门诊费用记录" & _
            " Where NO=[1] And 记录状态 IN(0,1,3) And 记录性质=1" & _
            "       And 收费类别 IN('5','6','7') And 执行部门ID is Not NULL" & _
            " Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then
        If rsTmp!收费类别 = "5" Then
            mlng西药房 = rsTmp!执行部门ID
        ElseIf rsTmp!收费类别 = "6" Then
            mlng成药房 = rsTmp!执行部门ID
        ElseIf rsTmp!收费类别 = "7" Then
            mlng中药房 = rsTmp!执行部门ID
        End If
    End If
    BillDrugDept = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDrugStuffBillDelData(ByRef lng结算序号 As Long, _
    ByRef strDelDrugFeeIDs As String, ByRef strDelDrugJosn As String, _
    ByRef strDelStuffFeeIDs As String, ByRef strDelStuffJosn As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取作废药品、材料的数据，针对历史退费异常
    '入参:
    '   lng结算序号 退费结算序号
    '出参:
    '   strDelDrugFeeIDs 作废药品涉及的费用ID
    '   strDelDrugJosn 作废药品需要的JSON入参
    '   strDelStuffFeeIDs 作废卫材涉及的费用ID
    '   strDelStuffJosn 作废卫材需要的JSON入参
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim strTemp As String

    On Error GoTo ErrHandler
    strDelDrugFeeIDs = "": strDelDrugJosn = ""
    strDelStuffFeeIDs = "": strDelStuffJosn = ""
    
    strSql = "Select a.Id, a.收费类别, b.数次" & _
            " From 门诊费用记录 A," & _
            "      (Select a.No, a.序号, -1 * Sum(Nvl(a.付数, 1) * a.数次) As 数次" & _
            "        From 门诊费用记录 A, 材料特性 B" & _
            "        Where a.收费细目id = b.材料id(+) And (a.收费类别 In ('5', '6', '7') Or a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1)" & _
            "              And a.结帐id In (Select 结帐id From 病人预交记录 Where 结算序号 = [1])" & _
            "        Group By a.No, a.序号" & _
            "        Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0) B" & _
            " Where a.记录性质 = 1 And a.记录状态 In (1, 3) And a.No = b.No And a.序号 = b.序号 " & _
            "   And Not Exists(Select 1 From 病人费用异常记录 c Where a.id = c.费用id And c.产生环节 = 1)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取未作废药品卫材", lng结算序号)
    If rsTemp.EOF Then GetDrugStuffBillDelData = True: Exit Function
    
    Do While Not rsTemp.EOF
        If InStr(",5,6,7,", "," & Nvl(rsTemp!收费类别) & ",") > 0 Then
            'Zl_Drugsvr_Delrecipebill
            '  --功能：药品处方销帐(或退费)
            '  --入参：Json_In:格式
            '  -- input
            '  --  item_list[]                更新数据列表[数组]
            '  --     rcpdtl_id           N 1 处方明细id,目前传入的费用ID
            '  --     chargeoffs_num      N 1 销帐数量
            strDelDrugFeeIDs = strDelDrugFeeIDs & "," & Nvl(rsTemp!ID)
            
            strTemp = ""
            strTemp = strTemp & "" & """rcpdtl_id"":" & Nvl(rsTemp!ID)
            strTemp = strTemp & "," & """chargeoffs_num"":" & Nvl(rsTemp!数次)
            strDelDrugJosn = strDelDrugJosn & ",{" & strTemp & "}"
        End If
        
        If Nvl(rsTemp!收费类别) = "4" Then
            'Zl_Stuffsvr_Delrecipebill
            '  --功能：卫材处方销帐(或退费)
            '  --入参：Json_In:格式
            '  -- input
            '  --  item_list[]                更新数据列表[数组]
            '  --     stuffdtl_id         N 1 处方明细id,目前传入的费用ID
            '  --     return_num          N 1 退料数量
            strDelStuffFeeIDs = strDelStuffFeeIDs & "," & Nvl(rsTemp!ID)
            
            strTemp = ""
            strTemp = strTemp & "" & """stuffdtl_id"":" & Nvl(rsTemp!ID)
            strTemp = strTemp & "," & """return_num"":" & Nvl(rsTemp!数次)
            strDelStuffJosn = strDelStuffJosn & ",{" & strTemp & "}"
        End If
        
        rsTemp.MoveNext
    Loop
    If strDelDrugFeeIDs <> "" Then strDelDrugFeeIDs = Mid(strDelDrugFeeIDs, 2)
    If strDelDrugJosn <> "" Then strDelDrugJosn = "{""input"":{""item_list"":[" & Mid(strDelDrugJosn, 2) & "]}}"
    If strDelStuffFeeIDs <> "" Then strDelStuffFeeIDs = Mid(strDelStuffFeeIDs, 2)
    If strDelStuffJosn <> "" Then strDelStuffJosn = "{""input"":{""item_list"":[" & Mid(strDelStuffJosn, 2) & "]}}"
    
    GetDrugStuffBillDelData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetNotSendDrugRec(ByVal bln未收费 As Boolean, _
    ByVal str开始时间 As String, ByVal str终止时间 As String, _
    ByRef rsData As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取未发药品记录
    '入参:
    '   str开始时间=查询开始时间:yyyy-mm-dd hh:mi:ss
    '   str终止时间=查询终止时间:yyyy-mm-dd hh:mi:ss
    '出参:
    '   rsData 字段：单据类型,处方单号,药房ID
    '       其中，单据类型:1-收费处方,2-记帐单处方,3-记帐表处方
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    If mobjService.zlDrugSvr_GetNotSendRec(1, bln未收费, "1,4", str开始时间, str终止时间, rsData, mlngModule) = False Then Exit Function
    GetNotSendDrugRec = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetNotSendStuffRec(ByVal bln未收费 As Boolean, _
    ByVal str开始时间 As String, ByVal str终止时间 As String, _
    ByRef rsData As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取未发料记录
    '入参:
    '   str开始时间=查询开始时间:yyyy-mm-dd hh:mi:ss
    '   str终止时间=查询终止时间:yyyy-mm-dd hh:mi:ss
    '出参:
    '   rsData 字段：单据类型,发料单号,库房ID
    '       其中，单据类型:1-收费发料单;2-记帐发料单;3-记帐表发料单
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    If mobjService.zlStuffSvr_GetNotSendRec(1, bln未收费, "1,4", str开始时间, str终止时间, rsData, mlngModule) = False Then Exit Function
    GetNotSendStuffRec = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

