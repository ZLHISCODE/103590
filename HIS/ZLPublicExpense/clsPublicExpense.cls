VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public zlBillingWarn As clsBillingWarn  '记帐报警相关
Public zlBillOperation As clsBillOperation '单据操作相关
Public zlDockExpense  As clsDockExpense '医嘱附费相关
Public zlInvoice     As clsInvoice   '发票相关
'***********************************************************************
'接口描述
'   zlShowMsgBox-消费框
'   zlSelectWholeItems-成套项目选择器
'***********************************************************************
 
Private Sub Class_Initialize()
    glngInstanceCount = glngInstanceCount + 1
    If zlBillingWarn Is Nothing Then Set zlBillingWarn = New clsBillingWarn
    If zlBillOperation Is Nothing Then Set zlBillOperation = New clsBillOperation
    If zlDockExpense Is Nothing Then Set zlDockExpense = New clsDockExpense
    If zlInvoice Is Nothing Then Set zlInvoice = New clsInvoice
End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not zlBillingWarn Is Nothing Then Set zlBillingWarn = Nothing
    If Not zlBillOperation Is Nothing Then Set zlBillOperation = Nothing
    If Not zlDockExpense Is Nothing Then Set zlDockExpense = Nothing
    If Not zlInvoice Is Nothing Then Set zlInvoice = Nothing
    
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    If zlGetComLib = False Then Exit Function
    
    Call GetUserInfo
    Call InitVar
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetPatiMoney(ByVal bytType As Byte, ByVal lng病人ID As Long, ByRef objPatiFee As clsPatiFeeinfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的相关费用信息
    '入参:bytType-0-门诊;1-住院
    '     lng病人ID-病人ID
     '出参:
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 16:45:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiMoney = GetPatiMoney(bytType, lng病人ID, objPatiFee)
End Function
Public Function zlGetFeeMoneyFromAdviceIDs(ByVal str医嘱IDs As String, _
    ByRef dblOut应收金额 As Double, ByRef dblOut实收金额 As Double) As Boolean
    '功能:根据医嘱IDs来获取应收和实收金额
    '入参:str医嘱IDs-医嘱ID,多个用逗号分离
    '出参:dblOut应收金额-应收金额
    '     dblOut实收金额-实收金额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 16:45:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetFeeMoneyFromAdviceIDs = GetFeeMoneyFromAdviceIDs(str医嘱IDs, dblOut应收金额, dblOut实收金额)
End Function

Public Function zlAdviceIsCharged(ByVal str医嘱IDs As String, _
    ByVal strNos As String, ByRef bytOutChargeStatus As Byte, Optional ByRef strOut未收医嘱IDs As String, _
    Optional ByRef bytOutBillType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断医嘱是否已经收费
    '入参:str医嘱IDs-医嘱ID,多个用逗号分离
    '     strNOs-单据号,多个用逗号分离
    '出参:bytOutChargeStatus-收费状态(0-部分收费,1-完全收费;2-部门收费)
    '     strOut未收医嘱IDs-返回未收费或未补审核的医嘱ID
    '     bytOutBillType:返回当前的单据类型(1-收费单;2-记帐单;3-收费和记帐都有)
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-26 09:48:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlAdviceIsCharged = AdviceIsCharged(str医嘱IDs, strNos, bytOutChargeStatus, strOut未收医嘱IDs, bytOutBillType)
End Function
Public Function zlBillExistNotBalance(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断收费单据是否存在未收费的
    '入参:strNOs:指定的单据号,允许多个,用逗号分离
    '出参:
    '返回:单据中存在未收费的,返回true,否则返回False
    '编制:冉俊明
    '日期:2016-08-25 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBillExistNotBalance = BillExistNotBalance(strNos)
End Function
Public Function zlGetBillChargeStatus(ByVal strNos As String, ByRef bytOutStatus As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费单据的计费状态
    '入参:strNOs:指定的单据号,允许多个,用逗号分离
    '出参:bytOutStatus:0-未收费;1-部分收费/退费;2-全部收费;3-全部退费
    '返回:获取成功,返回true,否则返回False(含未找到数据部分)
    '编制:刘兴洪
    '日期:2014-03-26 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetBillChargeStatus = GetBillChargeStatus(strNos, bytOutStatus)
End Function
Public Function zlGetBalanceStatus(ByVal strNos As String, ByRef bytOutStatus As Byte, _
    Optional bln门诊 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断记帐单是否已经结帐(只针对帐单)
    '入参:strNOs:指定的单据号,允许多个,用逗号分离
    '     bln门诊-门诊记帐单
    '出参:bytOutStatus:0-未结帐;1-部分结帐;2-全部结帐
    '返回:获取成功,返回true,否则返回False(含未找到数据部分)
    '编制:刘兴洪
    '日期:2014-03-26 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetBalanceStatus = GetBalanceStatus(strNos, bytOutStatus, bln门诊)
End Function
Public Function zlGetBalanceExpenseDetails(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lng结帐ID As Long, ByRef rsOutDetails As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定结帐的费用明细数据
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    lng结帐id -结帐ID
    '出参:rsOutDetails-结算数据(费用单号，收费类别、收费名称、收费数量、结帐金额，收费单价、计算单位、执行科室）
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-26 17:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetBalanceExpenseDetails = GetBalanceExpenseDetails(frmMain, lngModule, lng结帐ID, rsOutDetails)
End Function
Public Function zlGetBalanceInfor(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lng结帐ID As Long, ByRef rsOutBalance As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定结算数据
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    lng结帐id -结帐ID
    '出参:rsOutDetails-结算数据( 结算方式、结算金额、结算号码,医疗卡类别ID,消费卡,交易流水号,交易说明,刷卡卡号）
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-26 17:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetBalanceInfor = GetBalanceInfor(frmMain, lngModule, lng结帐ID, rsOutBalance)
End Function
Public Function zlIsExiseApplyNotAudit(ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
    ByRef blnOutExistApplyed As Boolean, ByRef rsOutMsgInfor As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断病人是存在销帐申请记录
    '入参:lng病人ID-病人ID
    '     lng主页ID-主页ID
    '出参:blnOutExistApplyed-true:表示存在销帐申请且未审核的销帐记录
    '     rsOutMsgInfor-返回申请记录(NO,项目名称,审核部门)
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-01 14:40:01
    '说明:根据病人ID和主页ID,判断该病人是否存在申请了的,但未审核的费用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim strInfo As String
    
    On Error GoTo errHandle
    strSql = "" & _
    "   Select distinct A.NO ,D.名称 项目名称,C.名称 审核科室  " & _
    "   From 住院费用记录 A, 病人费用销帐 B,部门表 C,收费项目目录 D" & vbNewLine & _
    "   Where A.病人ID=[1] And A.主页ID=[2] And A.Id = B.费用ID And b.状态=0  " & _
    "           And B.审核部门ID=C.ID And B.收费细目ID=D.ID" & _
    "   Order by A.NO,C.名称"
    Set rsOutMsgInfor = gobjDatabase.OpenSQLRecord(strSql, "检查是否存在销帐申请", lng病人ID, lng主页Id)
    blnOutExistApplyed = Not rsOutMsgInfor.EOF
    zlIsExiseApplyNotAudit = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function zlGetInvoiceNoToNOs(ByVal strInvioceNo As String, _
    ByRef strOutNos As String, Optional ByVal blnDataMoved As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据发票号获取该发票所对应的处方号(主要是针对收费单)
    '入参:strInvioceNo-发票号
    '     blnDataMoved-是否包含在历史空间表中查询
    '出参: strOutNOs-返回涉及的收费单号(多个用逗号分离)
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-12 15:59:32
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    On Error GoTo errHandle
    strOutNos = ""
    If gSysPara.byt票据分配规则 <> 0 Then
        strSql = "" & _
        "   Select distinct A.NO" & _
        "   From 票据打印明细 A" & _
        "   Where A.票号=[1] and 票种=1"
        If blnDataMoved Then
            strSql = strSql & vbCrLf & " UNION  ALL " & vbCrLf & _
             Replace(strSql, "票据打印明细", "H票据打印明细")
        End If
        strSql = strSql & " Order by NO"
    Else
        strSql = "" & _
         "   Select distinct NO  " & _
         "   From 票据打印内容 A," & _
         "           (   Select Max(M.打印ID) as 打印ID " & _
         "               From  票据使用明细 M  " & _
         "               Where M.票种=1 And M.性质=1 And M.号码=[1]  " & _
         "               Group by M.号码" & _
         "               )  Q" & _
         "   Where A.数据性质=1  And ID=Q.打印ID "
        If blnDataMoved Then
            strSql = strSql & vbCrLf & " UNION  ALL " & vbCrLf & _
             Replace(Replace(strSql, "票据打印内容", "H票据打印内容"), "票据使用明细", "H票据使用明细")
        End If
        strSql = strSql & " Order by NO"
    End If
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取对应发票的收费单据号", strInvioceNo)
    With rsTemp
        Do While Not .EOF
            strOutNos = strOutNos & "," & Nvl(!NO)
            .MoveNext
        Loop
        If strOutNos <> "" Then strOutNos = Mid(strOutNos, 2)
    End With
    zlGetInvoiceNoToNOs = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function zlGetPatiAuditRecord(ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
    ByRef rsOutAudit As ADODB.Recordset, Optional ByVal lng项目id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的费用审批项目
    '入参:lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     lng项目ID-项目ID(可选,可以不传)
    '出参:rsOutAudit-返回指定病人的费用审批项目信息
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-01 16:12:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    On Error GoTo errHandle
    
    strSql = "" & _
    "   Select 项目Id,使用限量,已用数量,使用限量-已用数量 可用数量  " & _
    "   From 病人审批项目  " & _
    "   Where 病人ID=[1] And 主页ID=[2]" & IIf(lng项目id <> 0, " And 项目ID=[3]", "")
    Set rsOutAudit = gobjDatabase.OpenSQLRecord(strSql, "mdlInExse", lng病人ID, lng主页Id, lng项目id)
    zlGetPatiAuditRecord = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function zlIsExiseNotAuditPriceBill(ByVal bytType As Byte, _
    ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
    ByVal str医嘱IDs As String, _
    ByRef blnOutExisePriceBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在未审核的划价单
    '入参:bytType-1-收费划价单;2-记帐划价单
    '     lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     str医嘱IDs-多个用逗号分离,允许为空,表示按病人来进行查询
    '出参:blnOutExisePriceBill-true:表示存在划价单;False:表示不存在划价单
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-01 16:20:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim varPara As Variant
    
    On Error GoTo errHandle
    If str医嘱IDs = "" Then
        If bytType = 1 Then
            strSql = "SELECT 1 From 门诊费用记录 where 记录性质=1 and 记录状态=0 and 病人ID=[1] and Rownum  <2"
        Else
            If lng主页Id = 0 Then
                strSql = "" & _
                "   SELECT 1 From 门诊费用记录 where  记录状态=0 And nvl( 记帐费用,0)=1 and 病人ID=[1]  and Rownum  <2 " & _
                "   UNION ALL " & _
                "   Select 1 From 住院费用记录 Where  记录状态=0 and  nvl( 记帐费用,0)=1  and 门诊标志<>2 and 病人ID=[1] and Rownum  <2"
            Else
                strSql = "SELECT 1 From 住院费用记录 where 记录状态=0 and  nvl( 记帐费用,0)=1   and 病人ID=[1] and 主页ID=[2] and Rownum  <2 "
            End If
        End If
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据病人ID判断是否存在划价单", lng病人ID, lng主页Id)
    Else
        If gobjCommFun.ActualLen(str医嘱IDs) >= 4000 Then
            If FromIDsBulidIngSQL(EM_Bulid_数字, str医嘱IDs, varPara, strSql, "医嘱ID") = False Then Exit Function
            strSql = "" & _
            " Select /*+ RULE */  1 " & _
            " From (With 医嘱数据 As (" & strSql & ") " & _
            "        SELECT 1 From 门诊费用记录 A,医嘱数据 B  Where a.医嘱序号=B.医嘱ID And A.记录状态=0 And nvl(A.记帐费用,0)=1 and Rownum  <2 " & _
            "        UNION ALL " & _
            "        SELECT 1 From 住院费用记录 A,医嘱数据 B  Where a.医嘱序号=B.医嘱ID And A.记录状态=0 And nvl(A.记帐费用,0)=1 and Rownum  <2 )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据医嘱ID判断是否存在划价单", varPara)
            
        Else
            strSql = "" & _
            " Select /*+ RULE */  1 " & _
            " From (With 医嘱数据 As (Select Column_Value As 医嘱id From Table(f_Num2list([1]))) " & _
            "        SELECT 1 From 门诊费用记录 A,医嘱数据 B  Where a.医嘱序号=B.医嘱ID And A.记录状态=0 And nvl(A.记帐费用,0)=1 and Rownum  <2   " & _
            "        UNION ALL " & _
            "        SELECT 1 From 住院费用记录 A,医嘱数据 B  Where a.医嘱序号=B.医嘱ID And A.记录状态=0 And nvl(A.记帐费用,0)=1 and Rownum  <2  )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据医嘱ID判断是否存在划价单", str医嘱IDs)
        End If
    End If
    blnOutExisePriceBill = Not rsTemp.EOF
    zlIsExiseNotAuditPriceBill = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetPriceGradeStartType() As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取价格等级的启用类型
    '返回: 0-未启用,1-只启用了站点,2-只启用了医疗付款方式,3-站点和医疗款方式都启用了
    '编制:刘兴洪
    '日期:2016-09-23 11:08:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPriceGradeStartType = GetPriceGradeStartType()
End Function

Public Function zlGetPriceGrade(ByVal str站点 As String, _
    ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
    Optional ByVal str医疗付款方式 As String, _
    Optional ByRef str药品价格等级_Out As String, _
    Optional ByRef str卫材价格等级_Out As String, _
    Optional ByRef str普通项目价格等级_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 根据医疗付款方式或站点，获取对应的价格等级
    '入参:str站点-登陆的站点，必须传入，传入NULL时，价格等级为返回空
    '     lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     str医疗付款方式:如果传入非空，则以传的医疗付款方式_In方式来提取价格等级;否则以病人ID_In或主页ID来获取对应的病人的医疗付款方式。
    
    '出参:str药品价格等级_out-返回药品价格等级
    '     str卫材价格等级_out-返回卫材价格等级
    '     str普通项目价格等级_out-返回普通收费项目价格等级
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-07-29 16:10:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPriceGrade = GetPriceGrade(str站点, lng病人ID, lng主页Id, str医疗付款方式, _
        str药品价格等级_Out, str卫材价格等级_Out, str普通项目价格等级_out)
End Function

Public Function zlGetRetailPrice(ByVal lng收费细目ID As Long, _
    ByVal str价格等级 As String, ByRef dbl零售价_out As Double, ByRef dbl未分解数_out As Double, _
    Optional ByVal lng库房ID As Long = 0, _
    Optional ByVal dbl数量 As Double = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 根据价格等级获取指定收费项目的零售价的相关信息
    '入参:lng收费细目id-收费细目ID
    '     str价格等级-收费价格等级
    '     lng库房id-库房ID（药品和卫生材料传入)
    '     dbl数量:当前出库数量(药品和卫生材料传入)。
    '出参:dbl零售价_out-返回零售价格
    '     dbl未分解数_out-针对药品和卫生材料有效，表示根据当前输入的出库数量（参数：dbl数量)进行分解时，未分解完的数量.
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-07-29 16:10:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetRetailPrice = GetRetailPrice(lng收费细目ID, str价格等级, dbl零售价_out, dbl未分解数_out, lng库房ID, dbl数量)
End Function

Public Function zlShowMsgBox(ByVal frmMain As Object, ByVal strInfo As String, Optional ByVal blnNoAsk As Boolean, Optional ByVal intType As Integer) As VbMsgBoxResult
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示消息框
    '入参:frmMain-调用的主窗体
    '     strInfo=提示信息,需要自已处理换行,可用"^"表示回车,">"表示缩进
    '     intType=消息框类型=0(缺省)=MsgBox类型,1-皮试类型
    '     blnNoAsk="intType=0"时有效，表示是否只显示一个确定按钮,不以询问方式显示是和否。
    '返回:
    '    intType=0：vbIgnore=是且不再提示,vbCancel=否且不再提示,vbYes=是,vbNo=否
    '    intType=1：vbYes=阳性,vbNo=阴性,vbCancel=取消
    '编制:刘兴洪
    '日期:2017-11-08 11:17:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
  zlShowMsgBox = ShowMsgBox_Custom(frmMain, strInfo, blnNoAsk, intType)
End Function
Public Function zlSelectWholeItems(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
     ByRef rsOutSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示成套项目选择器(选择器入口)
    '入参:lngModule-模块号
    '       strPrivs-权限串
    '出参:rsOutSel-成功时,返回选择的成套项目(有字段:细目ID,编码,名称,序号,从属父号,执行科室....)
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-08 16:22:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSelectWholeItems = SelectWholeItems(frmMain, lngModule, strPrivs, rsOutSel)
End Function

Public Function zlPatiIdentify(ByVal lngModlue As Long, ByVal frmMain As Object, ByVal lng病人ID As Long, ByVal curMoney As Currency, _
    Optional ByVal bln退费 As Boolean = False, Optional ByVal bytDepositShowMode As Byte = 0, Optional ByVal lngDefaultCardTypeID As Long = 0, _
    Optional ByVal blnFamilyMoney As Boolean, Optional ByVal blnOlnyFamilyIDs As Boolean, Optional strFamilyPatiIDs_Out As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行刷卡验证
    '入参:lngModlue-模块号
    '     dblMoney-金额
    '     lng病人ID-病人ID
    '     bln退费-当前是否退费操作
    '     bytDepositShowMode- 预交显示方式(0-余额汇总显示;1-只显示门诊余额;2-只显示住院余额)
    '     lngDefaultCardTypeID-缺省的刷卡类别
    '     blnFamilyMoney-是否读取家属预交余额
    '     blnOlnyFamilyIDs-true:不验卡，只读取家属IDs;False-需要读取卡验卡
    '出参:strFamilyPatiIDs-病人家属ID,多个用逗号分隔，79868
    '返回:身份验证成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-10-24 14:55:59
    '说明:
    '   一、消费验卡输入密码规则（bln退费=false时):
    '       1.不进行刷卡验证,直接返回True
    '       2.门诊消费时，需要进行刷卡验证，同时需要输入密码（无密码时,光标要经过密码框)
    '       3.门诊消费时设置密码的，则必须刷卡验卡且输入密码，无密码时,不需要验卡输入密码
    '       4.表示病人在消费N元内必须刷卡,不必输入密码即可支付;否则必须输入密码(无密码时，光标要经过密码框)
    '  二、退费验卡（bln退费=true时):
    '       1.不进行刷卡控制，直接返回true
    '       2.门诊消费退费时需要刷卡验证,同时需要输入密码（无密码时,光标要经过密码框)
    '       3.门诊消费退费时设置密码的，则必须刷卡验证,无密码时,不需要验卡输入密码
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiIdentify = PatiIdentify(lngModlue, frmMain, lng病人ID, curMoney, bln退费, bytDepositShowMode, lngDefaultCardTypeID, _
                                               blnFamilyMoney, blnOlnyFamilyIDs, strFamilyPatiIDs_Out)
End Function




