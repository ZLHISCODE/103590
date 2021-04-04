Attribute VB_Name = "mdlParameter"
Option Explicit
Private mstr用户名 As String
Private mstr机器名 As String

Public Sub UpdateParameters()
'功能：对原本机的注册表参数值进行升级处理
    Dim rsTmp As New ADODB.Recordset
    Dim rsSys As New ADODB.Recordset
    Dim rsUpgrade As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select User as 用户名,SYS_CONTEXT('USERENV','TERMINAL') as 机器名 From Dual"
    Call zlDataBase.OpenRecordset(rsTmp, strSQL, "UpdateParameters")
    mstr用户名 = rsTmp!用户名: mstr机器名 = rsTmp!机器名
    
    strSQL = "Select Trunc(编号/100) as 系统 From zlSystems Where 版本号 Like '10.%'"
    Call zlDataBase.OpenRecordset(rsSys, strSQL, "UpdateParameters")
    
    strSQL = "Select Trunc(系统/100) as 系统 From zlUpgrade" & _
        " Where 系统 Is Not Null And 原始版本 Like '10.%'" & _
        " And 原始版本>='10.24.0' And Substr(目标版本,1,5)>Substr(原始版本,1,5)"
    Call zlDataBase.OpenRecordset(rsUpgrade, strSQL, "UpdateParameters")
    
    On Error GoTo 0
    
    '处理标准版的参数值升级
    '-----------------------------------------------------------------
    rsSys.Filter = "系统=1": rsUpgrade.Filter = "系统=1"
    If Not rsSys.EOF And rsUpgrade.EOF Then
        '挂号管理
        Call UpdateParameterValue(rsSys!系统, 1111, "公共模块\zl9RegEvent", "缺省付款方式", "缺省付款方式")
        Call UpdateParameterValue(rsSys!系统, 1111, "公共模块\zl9RegEvent", "缺省费别", "缺省费别")
        Call UpdateParameterValue(rsSys!系统, 1111, "公共模块\zl9RegEvent", "缺省性别", "缺省性别")
        Call UpdateParameterValue(rsSys!系统, 1111, "公共模块\zl9RegEvent", "缺省结算方式", "缺省结算方式")
        Call UpdateParameterValue(rsSys!系统, 1111, "公共模块\zl9RegEvent", "挂号科室", "挂号科室")
        Call UpdateParameterValue(rsSys!系统, 1111, "公共模块\zl9RegEvent", "共用挂号票据批次", "共用挂号票据批次")
        Call UpdateParameterValue(rsSys!系统, 1111, "公共模块\zl9RegEvent", "共用就诊卡批次", "共用就诊卡批次")
        Call UpdateParameterValue(rsSys!系统, 1111, "私有模块\" & mstr用户名 & "\zl9RegEvent", "当前挂号票据号", "当前挂号票据号")
        Call UpdateParameterValue(rsSys!系统, 1111, "私有模块\" & mstr用户名 & "\zl9RegEvent\frmRegist", "刷新方式", "刷新方式")
        '门诊分诊管理
        Call UpdateParameterValue(rsSys!系统, 1113, "公共模块\zl9RegEvent", "分诊科室", "分诊科室")
        '门诊收费
        Call UpdateParameterValue(rsSys!系统, 1121, "公共模块\zl9OutExse", "收费类别", "收费类别")
        Call UpdateParameterValue(rsSys!系统, 1121, "私有模块\" & mstr用户名 & "\zl9OutExse", "缺省费别", "缺省费别")
        Call UpdateParameterValue(rsSys!系统, 1121, "私有模块\" & mstr用户名 & "\zl9OutExse\frmManageCharge", "刷新方式", "刷新方式")
        Call UpdateParameterValue(rsSys!系统, 1121, "私有模块\" & mstr用户名 & "\zl9OutExse", "中药自动输入", "中药自动输入")
        Call UpdateParameterValue(rsSys!系统, 1121, "私有模块\" & mstr用户名 & "\zl9OutExse", "中药自动输入长度", "中药自动输入长度")
        Call UpdateParameterValue(rsSys!系统, 1121, "私有模块\" & mstr用户名 & "\zl9OutExse", "缺省结算方式", "缺省结算方式")
        Call UpdateParameterValue(rsSys!系统, 1121, "公共模块\zl9OutExse", "共用收费票据批次", "共用收费票据批次")
        Call UpdateParameterValue(rsSys!系统, 1121, "公共模块\zl9OutExse", "挂号共用收费票据", "挂号共用收费票据")
        Call UpdateParameterValue(rsSys!系统, 1121, "公共模块\zl9OutExse", "手工报价", "手工报价")
        Call UpdateParameterValue(rsSys!系统, 1121, "公共模块\zl9OutExse", "LED显示收费明细", "LED显示收费明细")
        Call UpdateParameterValue(rsSys!系统, 1121, "公共模块\zl9OutExse", "LED显示欢迎信息", "LED显示欢迎信息")
        Call UpdateParameterValue(rsSys!系统, 1121, "私有模块\" & mstr用户名 & "\zl9OutExse", "当前收费票据号", "当前收费票据号")
        Call UpdateParameterValue(rsSys!系统, 1121, "私有模块\" & mstr用户名 & "\zl9OutExse", "退费号码输入模式", "退费号码输入模式")
        
        '门诊划价
        Call UpdateParameterValue(rsSys!系统, 1120, "公共模块\zl9OutExse", "收费类别", "收费类别")
        Call UpdateParameterValue(rsSys!系统, 1120, "私有模块\" & mstr用户名 & "\zl9OutExse", "缺省费别", "缺省费别")
        Call UpdateParameterValue(rsSys!系统, 1120, "私有模块\" & mstr用户名 & "\zl9OutExse\frmManagePrice", "刷新方式", "刷新方式")
        Call UpdateParameterValue(rsSys!系统, 1120, "私有模块\" & mstr用户名 & "\zl9OutExse", "中药自动输入", "中药自动输入")
        Call UpdateParameterValue(rsSys!系统, 1120, "私有模块\" & mstr用户名 & "\zl9OutExse", "中药自动输入长度", "中药自动输入长度")
        
        '门诊记帐
        Call UpdateParameterValue(rsSys!系统, 1122, "公共模块\zl9OutExse", "收费类别", "收费类别")
        Call UpdateParameterValue(rsSys!系统, 1122, "私有模块\" & mstr用户名 & "\zl9OutExse\frmManageBilling", "刷新方式", "刷新方式")
        Call UpdateParameterValue(rsSys!系统, 1122, "私有模块\" & mstr用户名 & "\zl9OutExse\frmManageBilling\TabStrip", "页面", "页面")
        
        '住院记帐管理
        Call UpdateParameterValue(rsSys!系统, 1133, "私有模块\" & mstr用户名 & "\zl9InExse\frmManageBilling", "刷新方式", "刷新方式")
        Call UpdateParameterValue(rsSys!系统, 1133, "私有模块\" & mstr用户名 & "\zl9InExse\frmManageBilling\TabStrip", "页面", "页面")
        
        '科室分散记帐
        Call UpdateParameterValue(rsSys!系统, 1134, "私有模块\" & mstr用户名 & "\zlInExse\frmDeptBilling", "刷新方式", "刷新方式")
        Call UpdateParameterValue(rsSys!系统, 1134, "私有模块\" & mstr用户名 & "\zl9InExse\frmDeptBilling\TabStrip", "页面", "页面")
        Call UpdateParameterValue(rsSys!系统, 1134, "私有模块\" & mstr用户名 & "\zl9InExse\frmDeptBilling", "显示病人方式", "显示病人方式")
        
        '医技科室记帐
        Call UpdateParameterValue(rsSys!系统, 1135, "私有模块\" & mstr用户名 & "\zl9InExse\frmTechnoBilling", "刷新方式", "刷新方式")
        
        '执行登记管理
        Call UpdateParameterValue(rsSys!系统, 1142, "公共模块\zl9InExse", "医技病人来源", "医技病人来源")
        Call UpdateParameterValue(rsSys!系统, 1142, "公共模块\zl9InExse", "医技门诊单据类型", "医技门诊单据类型")
        Call UpdateParameterValue(rsSys!系统, 1142, "公共模块\zl9InExse", "医技住院单据类型", "医技住院单据类型")
        Call UpdateParameterValue(rsSys!系统, 1142, "公共模块\zl9InExse", "医技体检单据类型", "医技体检单据类型")
        Call UpdateParameterValue(rsSys!系统, 1142, "公共模块\zl9InExse", "医技执行类别", "医技执行类别")
        Call UpdateParameterValue(rsSys!系统, 1142, "私有模块\" & mstr用户名 & "\zl9InExse", "显示单据头", "显示单据头", True)
        Call UpdateParameterValue(rsSys!系统, 1142, "公共模块\zl9InExse", "医技主从项目同时选择", "主从项目同时选择")
        
        '住院记帐操作
        Call UpdateParameterValue(rsSys!系统, 1150, "私有模块\" & mstr用户名 & "\zl9InExse\frmPatiSelect", "显示病人方式", "显示病人方式")
        Call UpdateParameterValue(rsSys!系统, 1150, "私有模块\" & mstr用户名 & "\zl9InExse", "显示在院病人", "显示在院病人")
        Call UpdateParameterValue(rsSys!系统, 1150, "私有模块\" & mstr用户名 & "\zl9InExse", "显示预出院病人", "显示预出院病人")
        Call UpdateParameterValue(rsSys!系统, 1150, "私有模块\" & mstr用户名 & "\zl9InExse", "显示出院病人", "显示出院病人")
        Call UpdateParameterValue(rsSys!系统, 1150, "私有模块\" & mstr用户名 & "\zl9InExse\frmReCharge", "费用开始时间", "费用开始时间")
        Call UpdateParameterValue(rsSys!系统, 1150, "私有模块\" & mstr用户名 & "\zl9InExse\frmReCharge", "忽略期间", "忽略期间")
        Call UpdateParameterValue(rsSys!系统, 1150, "私有模块\" & mstr用户名 & "\zl9InExse\frmReCharge", "审核开始时间", "审核开始时间")
        Call UpdateParameterValue(rsSys!系统, 1150, "私有模块\" & mstr用户名 & "\zl9InExse", "中药自动输入", "中药自动输入")
        Call UpdateParameterValue(rsSys!系统, 1150, "私有模块\" & mstr用户名 & "\zl9InExse", "中药自动输入长度", "中药自动输入长度")
        Call UpdateParameterValue(rsSys!系统, 1150, "公共模块\zl9InExse", "门诊留观病人记帐", "门诊留观病人记帐")
        Call UpdateParameterValue(rsSys!系统, 1150, "公共模块\zl9InExse", "住院留观病人记帐", "住院留观病人记帐")
        Call UpdateParameterValue(rsSys!系统, 1150, "公共模块\zl9InExse", "收费类别", "收费类别")
        
        '病人结帐管理
        Call UpdateParameterValue(rsSys!系统, 1137, "私有模块\" & mstr用户名 & "\zl9InExse", "结帐票据类型", "结帐票据类型")
        Call UpdateParameterValue(rsSys!系统, 1137, "公共模块\zl9InExse", "共用结帐票据批次", "共用结帐票据批次")
        Call UpdateParameterValue(rsSys!系统, 1137, "公共模块\zl9InExse", "LED显示欢迎信息", "LED显示欢迎信息")
        Call UpdateParameterValue(rsSys!系统, 1137, "私有模块\" & mstr用户名 & "\zl9InExse", "当前结帐票据号", "当前结帐票据号")
        Call UpdateParameterValue(rsSys!系统, 1137, "私有模块\" & mstr用户名 & "\zl9InExse\frmManageBalance", "刷新方式", "刷新方式")
        Call UpdateParameterValue(rsSys!系统, 1137, "私有模块\" & mstr用户名 & "\zl9InExse\frmManageDue\TabStrip", "页面", "病人应收款页面")
        Call UpdateParameterValue(rsSys!系统, 1137, "私有模块\" & mstr用户名 & "\zl9InExse\frmPatiSelect", "显示病人方式", "显示病人方式")
        Call UpdateParameterValue(rsSys!系统, 1137, "私有模块\" & mstr用户名 & "\zl9InExse", "显示在院病人", "显示在院病人")
        Call UpdateParameterValue(rsSys!系统, 1137, "私有模块\" & mstr用户名 & "\zl9InExse", "显示预出院病人", "显示预出院病人")
        Call UpdateParameterValue(rsSys!系统, 1137, "私有模块\" & mstr用户名 & "\zl9InExse", "显示出院病人", "显示出院病人")
        Call UpdateParameterValue(rsSys!系统, 1137, "私有模块\" & mstr用户名 & "\zl9InExse", "显示结清病人", "显示结清病人")
        Call UpdateParameterValue(rsSys!系统, 1137, "私有模块\" & mstr用户名 & "\zl9InExse", "默认出院结帐", "默认出院结帐", True)
        
        '一日清单管理
        Call UpdateParameterValue(rsSys!系统, 1141, "私有模块\" & mstr用户名 & "\zl9InExse", "一日清单－病人病区模式", "病人病区模式")
        Call UpdateParameterValue(rsSys!系统, 1141, "私有模块\" & mstr用户名 & "\zl9InExse", "一日清单－结束时间", "结束时间")
        Call UpdateParameterValue(rsSys!系统, 1141, "私有模块\" & mstr用户名 & "\zl9InExse", "一日清单－结束间隔", "结束间隔")
        Call UpdateParameterValue(rsSys!系统, 1141, "私有模块\" & mstr用户名 & "\zl9InExse", "一日清单－开始时间", "开始时间")
        Call UpdateParameterValue(rsSys!系统, 1141, "私有模块\" & mstr用户名 & "\zl9InExse", "一日清单－开始间隔", "开始间隔")
        Call UpdateParameterValue(rsSys!系统, 1141, "私有模块\" & mstr用户名 & "\zl9InExse", "一日清单－非医保病人", "非医保病人")
        Call UpdateParameterValue(rsSys!系统, 1141, "私有模块\" & mstr用户名 & "\zl9InExse", "一日清单－医保病人", "医保病人")
        Call UpdateParameterValue(rsSys!系统, 1141, "私有模块\" & mstr用户名 & "\zl9InExse", "一日清单－在院病人", "在院病人")
        Call UpdateParameterValue(rsSys!系统, 1141, "私有模块\" & mstr用户名 & "\zl9InExse", "一日清单－出院病人", "出院病人")
        
        '病人费用查询
        Call UpdateParameterValue(rsSys!系统, 1139, "私有模块\" & mstr用户名 & "\zl9InExse", "清单比例", "清单比例")
        Call UpdateParameterValue(rsSys!系统, 1139, "私有模块\" & mstr用户名 & "\zl9InExse", "ViewDate", "费用时间类型", True)
        Call UpdateParameterValue(rsSys!系统, 1139, "私有模块\" & mstr用户名 & "\zl9InExse", "ViewCancel状态", "显示结帐作废", True)
        Call UpdateParameterValue(rsSys!系统, 1139, "私有模块\" & mstr用户名 & "\zl9InExse", "ViewZero状态", "显示零费用", True)
        Call UpdateParameterValue(rsSys!系统, 1139, "公共模块\zl9InExse", "显示体检费用", "显示体检费用", True)
        Call UpdateParameterValue(rsSys!系统, 1139, "公共模块\zl9InExse", "分科模式", "分科模式")
        Call UpdateParameterValue(rsSys!系统, 1139, "公共模块\zl9InExse", "分类模式", "分类模式")
        Call UpdateParameterValue(rsSys!系统, 1139, "私有模块\" & mstr用户名 & "\zl9InExse", "病人状态", "病人状态")
        Call UpdateParameterValue(rsSys!系统, 1139, "私有模块\" & mstr用户名 & "\zl9InExse", "欠费查询-间隔天数", "间隔天数")
        Call UpdateParameterValue(rsSys!系统, 1139, "私有模块\" & mstr用户名 & "\zl9InExse", "ViewOwe状态", "仅显未结清病人", True)
        Call UpdateParameterValue(rsSys!系统, 1139, "私有模块\" & mstr用户名 & "\zl9InExse", "ViewUnAudit状态", "仅显未审核病人", True)
        Call UpdateParameterValue(rsSys!系统, 1139, "私有模块\" & mstr用户名 & "\zl9InExse", "欠费查询-单次显示", "单次显示")
        
        '票据使用监控
        Call UpdateParameterValue(rsSys!系统, 1501, "私有模块\" & mstr用户名 & "\zL9CashBill\frmBillSupervise\Menu", "mnuViewAll状态", "显示所有领用记录", True)
        Call UpdateParameterValue(rsSys!系统, 1501, "私有模块\" & mstr用户名 & "\zL9CashBill\frmBillSupervise\Menu", "查看核对信息", "查看核对信息", True)
        
        '收费财务监控
        Call UpdateParameterValue(rsSys!系统, 1500, "私有模块\" & mstr用户名 & "\zL9CashBill\frmCashSupervise\Menu", "mnuViewAll状态", "显示所有收款员", True)
        
    
        '1260-门诊医生站
        Call UpdateParameterValue(rsSys!系统, 1260, "公共模块\zl9CISJob", "本地诊室", "本地诊室")
        Call UpdateParameterValue(rsSys!系统, 1260, "公共模块\zl9CISJob", "本机门诊科室", "本机门诊科室")
        Call UpdateParameterValue(rsSys!系统, 1260, "公共模块\zl9CISJob", "接诊范围", "接诊范围")
        Call UpdateParameterValue(rsSys!系统, 1260, "公共模块\zl9CISJob", "接诊科室", "接诊科室")
        Call UpdateParameterValue(rsSys!系统, 1260, "私有模块\" & mstr用户名 & "\zl9CISJob", "接诊医生", "接诊医生")
        Call UpdateParameterValue(rsSys!系统, 1260, "私有模块\" & mstr用户名 & "\zl9CISJob", "已诊病人结束间隔", "已诊病人结束间隔")
        Call UpdateParameterValue(rsSys!系统, 1260, "私有模块\" & mstr用户名 & "\zl9CISJob", "已诊病人开始间隔", "已诊病人开始间隔")
        Call UpdateParameterValue(rsSys!系统, 1260, "私有模块\" & mstr用户名 & "\zl9CISJob\frmOutDoctorStation", "医护功能", "医护功能")
        
        '1261-住院医生站
        Call UpdateParameterValue(rsSys!系统, 1261, "私有模块\" & mstr用户名 & "\zl9CISJob\frmAuditResponse", "反馈条件-随机抽查", "随机抽查反馈")
        Call UpdateParameterValue(rsSys!系统, 1261, "私有模块\" & mstr用户名 & "\zl9CISJob\frmAuditResponse", "反馈条件-提交审查", "提交审查反馈")
        Call UpdateParameterValue(rsSys!系统, 1261, "私有模块\" & mstr用户名 & "\zl9CISJob\frmInDoctorStation", "医护功能", "医护功能")
        
        '1262-住院护士站
        Call UpdateParameterValue(rsSys!系统, 1262, "私有模块\" & mstr用户名 & "\zl9CISJob\frmAuditResponse", "反馈条件-随机抽查", "随机抽查反馈")
        Call UpdateParameterValue(rsSys!系统, 1262, "私有模块\" & mstr用户名 & "\zl9CISJob\frmAuditResponse", "反馈条件-提交审查", "提交审查反馈")
        Call UpdateParameterValue(rsSys!系统, 1262, "私有模块\" & mstr用户名 & "\zl9CISJob\frmInNurseStation", "Filter当前病况", "当前病况过滤")
        Call UpdateParameterValue(rsSys!系统, 1262, "私有模块\" & mstr用户名 & "\zl9CISJob\frmInNurseStation", "Filter护理等级", "护理等级过滤")
        Call UpdateParameterValue(rsSys!系统, 1262, "私有模块\" & mstr用户名 & "\zl9CISJob\frmInNurseStation", "医护功能", "医护功能")
        
        '1263-医技工作站
        Call UpdateParameterValue(rsSys!系统, 1263, "公共模块\zl9CISJob", "记录执行情况", "记录执行情况")
        Call UpdateParameterValue(rsSys!系统, 1263, "私有模块\" & mstr用户名 & "\zl9CISJob\frmTechnicStation", "医护功能", "医护功能")
        '这个参数特殊处理
        Set rsTmp = New ADODB.Recordset
        strSQL = "Select A.部门ID From 部门人员 A,上机人员表 B Where A.人员ID=B.人员ID And B.用户名=User"
        Call zlDataBase.OpenRecordset(rsTmp, strSQL, "UpdateParameters")
        Do While Not rsTmp.EOF
            Call UpdateParameterValue(rsSys!系统, 1263, "公共模块\zl9CISJob\科室" & rsTmp!部门ID, "执行间范围", "执行间范围")
            rsTmp.MoveNext
        Loop

        '1252-门诊医嘱下达
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel", "门诊缺省成药房", "门诊缺省成药房")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel", "门诊缺省发料部门", "门诊缺省发料部门")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel", "门诊缺省西药房", "门诊缺省西药房")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel", "门诊缺省中药房", "门诊缺省中药房")
        Call UpdateParameterValue(rsSys!系统, 1252, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockOutAdvice", "FilterAutoHide", "过滤条件自动隐藏")
        Call UpdateParameterValue(rsSys!系统, 1252, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockOutAdvice", "Filter病人婴儿", "病人婴儿过滤")
        Call UpdateParameterValue(rsSys!系统, 1252, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockOutAdvice", "Filter科内医嘱", "科内医嘱过滤")
        Call UpdateParameterValue(rsSys!系统, 1252, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockOutAdvice", "Filter需要报告", "需要报告过滤")
        Call UpdateParameterValue(rsSys!系统, 1252, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockOutAdvice", "Filter医嘱状态", "医嘱状态过滤")
        Call UpdateParameterValue(rsSys!系统, 1252, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockOutAdvice", "医嘱子列表", "医嘱子列表")
        Call UpdateParameterValue(rsSys!系统, 1252, "私有模块\" & mstr用户名 & "\zlCISKernel\frmLisView", "隐藏检验图形", "隐藏检验图形")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel\frmLisRptGeneral", "查看中文", "查看中文")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel\frmLisRptGeneral", "查看标志", "查看标志")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel\frmLisRptGeneral", "查看单位", "查看单位")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel\frmLisRptGeneral", "查看参考", "查看参考")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel\frmLisRptGeneral", "查看酶标", "查看酶标")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel\frmLisRptGeneral", "查看备注", "查看备注")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel\frmLisRptGeneral", "查看诊断", "查看诊断")
        Call UpdateParameterValue(rsSys!系统, 1252, "公共模块\zlCISKernel\frmLisRptMicrobiology", "上次结果", "上次结果")
        
        '1253-住院医嘱下达
        Call UpdateParameterValue(rsSys!系统, 1253, "公共模块\zlCISKernel", "住院缺省成药房", "住院缺省成药房")
        Call UpdateParameterValue(rsSys!系统, 1253, "公共模块\zlCISKernel", "住院缺省发料部门", "住院缺省发料部门")
        Call UpdateParameterValue(rsSys!系统, 1253, "公共模块\zlCISKernel", "住院缺省西药房", "住院缺省西药房")
        Call UpdateParameterValue(rsSys!系统, 1253, "公共模块\zlCISKernel", "住院缺省中药房", "住院缺省中药房")
        Call UpdateParameterValue(rsSys!系统, 1253, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockInAdvice", "FilterAutoHide", "过滤条件自动隐藏")
        Call UpdateParameterValue(rsSys!系统, 1253, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockInAdvice", "Filter病人婴儿", "病人婴儿过滤")
        Call UpdateParameterValue(rsSys!系统, 1253, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockInAdvice", "Filter科内医嘱", "科内医嘱过滤")
        Call UpdateParameterValue(rsSys!系统, 1253, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockInAdvice", "Filter需要报告", "需要报告过滤")
        Call UpdateParameterValue(rsSys!系统, 1253, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockInAdvice", "Filter医嘱期效", "医嘱期效过滤")
        Call UpdateParameterValue(rsSys!系统, 1253, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockInAdvice", "Filter医嘱状态", "医嘱状态过滤")
        Call UpdateParameterValue(rsSys!系统, 1253, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockInAdvice", "Filter重整医嘱", "重整医嘱过滤")
        Call UpdateParameterValue(rsSys!系统, 1253, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDockInAdvice", "医嘱子列表", "医嘱子列表")
        
        '1254-住院医嘱发送
        Call UpdateParameterValue(rsSys!系统, 1254, "公共模块\zlCISKernel", "缺省留存比例", "缺省留存比例")
        Call UpdateParameterValue(rsSys!系统, 1254, "公共模块\zlCISKernel", "缺省留存计算", "缺省留存计算")
        Call UpdateParameterValue(rsSys!系统, 1254, "公共模块\zlCISKernel", "缺省留存药房", "缺省留存药房")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceOperateCond", "上次开始暂停", "上次开始暂停")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceReport", "常用报表病人", "常用报表病人")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceReport", "常用报表结束间隔", "常用报表结束间隔")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceReport", "常用报表结束时点", "常用报表结束时点")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceReport", "常用报表开始间隔", "常用报表开始间隔")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceReport", "常用报表开始时点", "常用报表开始时点")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceReport", "常用报表期效", "常用报表期效")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceRollSendCond", "超期收回病人", "超期收回病人")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendDrugCond", "限制结束时间", "药嘱发送限制结束时间")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendDrugCond", "药嘱发送病人", "药嘱发送病人")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendDrugCond", "药嘱给药途径", "药嘱发送给药途径")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendDrugCond", "药嘱结束时点", "药嘱发送结束时点")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendDrugCond", "药嘱结束时间", "药嘱发送结束时间")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendDrugCond", "药嘱时间间隔", "药嘱发送时间间隔")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendDrugCond", "药嘱药房置换", "药嘱发送药房置换")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendDrugCond", "药嘱医嘱期效", "药嘱发送医嘱期效")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendOtherCond", "非药发送病人", "其他发送病人")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendOtherCond", "非药结束时点", "其他发送结束时点")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendOtherCond", "非药结束时间", "其他发送结束时间")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendOtherCond", "非药时间间隔", "其他发送时间间隔")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendOtherCond", "非药医嘱期效", "其他发送医嘱期效")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmAdviceSendOtherCond", "非药诊疗类别", "其他发送诊疗类别")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDrugSendQueryCond", "退药查询间隔", "退药查询间隔")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDrugSendQueryCond", "药疗查询病人", "药疗查询病人")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDrugSendQueryCond", "药疗查询出院病人", "药疗查询出院病人")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDrugSendQueryCond", "药疗查询间隔", "药疗查询间隔")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDrugSendQueryCond", "药疗查询期效", "药疗查询期效")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDrugSendQueryCond", "药疗查询药房", "药疗查询药房")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDrugSendQueryCond", "药疗查询状态", "药疗查询状态")
        Call UpdateParameterValue(rsSys!系统, 1254, "私有模块\" & mstr用户名 & "\zlCISKernel\frmDrugSendQueryCond", "药嘱给药途径", "药疗查询给药途径")
        
        '1255-护理记录管理
        Call UpdateParameterValue(rsSys!系统, 1255, "公共模块\zlRichEPR\体温单打印选项", "不打印脉搏短绌图形", "不打印脉搏短绌图形")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\frmCaseTendBodyPrintSet", "连续打印", "连续打印")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\frmCaseTendBodyPrintSet", "打印页号", "打印页号")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\frmCaseTendBodyPrintSet", "起始页号", "起始页号")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\体温单打印选项", "打印周数", "打印周数")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\frmCaseTendSign", "chkEsign", "护理数字签名")
        
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\打印设置", "打印机", "体温单打印机")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\打印设置", "纸张", "体温单纸张")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\打印设置", "宽度", "体温单宽度")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\打印设置", "高度", "体温单高度")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\打印设置", "纸向", "体温单纸向")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\打印设置", "进纸", "体温单进纸")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\打印设置", "左边距", "体温单左边距")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\打印设置", "右边距", "体温单右边距")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\打印设置", "上边距", "体温单上边距")
        Call UpdateParameterValue(rsSys!系统, 1255, "私有模块\" & mstr用户名 & "\zlRichEPR\打印设置", "下边距", "体温单下边距")
        
        '1257-医嘱附费管理
        Call UpdateParameterValue(rsSys!系统, 1257, "公共模块\zlCISKernel", "门诊缺省成药房", "门诊缺省成药房")
        Call UpdateParameterValue(rsSys!系统, 1257, "公共模块\zlCISKernel", "门诊缺省西药房", "门诊缺省西药房")
        Call UpdateParameterValue(rsSys!系统, 1257, "公共模块\zlCISKernel", "门诊缺省中药房", "门诊缺省中药房")
        Call UpdateParameterValue(rsSys!系统, 1257, "公共模块\zlCISKernel", "收费类别", "收费类别")
        Call UpdateParameterValue(rsSys!系统, 1257, "公共模块\zlCISKernel", "住院缺省成药房", "住院缺省成药房")
        Call UpdateParameterValue(rsSys!系统, 1257, "公共模块\zlCISKernel", "住院缺省西药房", "住院缺省西药房")
        Call UpdateParameterValue(rsSys!系统, 1257, "公共模块\zlCISKernel", "住院缺省中药房", "住院缺省中药房")

        '1264-门诊输液排队
        Call UpdateParameterValue(rsSys!系统, 1264, "公共模块\zl9Transfusion", "显示单据种类", "显示单据种类")
        
        '处理的应付款系统的相关模块
        '刘兴宏
        '付款管理
        Call UpdateParameterValue(rsSys!系统, 1323, "私有模块\" & mstr用户名 & "\zl9Due\付款管理", "表头", "付款表头列表")
        Call UpdateParameterValue(rsSys!系统, 1323, "私有模块\" & mstr用户名 & "\zl9Due\付款管理", "付款明细", "付款明细列表")
        Call UpdateParameterValue(rsSys!系统, 1323, "私有模块\" & mstr用户名 & "\zl9Due\付款管理", "付款信息", "付款方式列表")
        '应付查询
        Call UpdateParameterValue(rsSys!系统, 1324, "私有模块\" & mstr用户名 & "\zl9Due\应付款查询", "单位ID", "最后选择单位ID")
        Call UpdateParameterValue(rsSys!系统, 1324, "私有模块\" & mstr用户名 & "\zl9Due\应付款查询", "供应商余额", "余额信息列表")
        Call UpdateParameterValue(rsSys!系统, 1324, "私有模块\" & mstr用户名 & "\zl9Due\应付款查询", "应付款查询-付款明细", "付款明细列表")
        Call UpdateParameterValue(rsSys!系统, 1324, "私有模块\" & mstr用户名 & "\zl9Due\应付款查询", "应付款查询-已付明细", "已付明细列表")
        Call UpdateParameterValue(rsSys!系统, 1324, "私有模块\" & mstr用户名 & "\zl9Due\应付款查询", "应付款查询-未付明细", "未付明细列表")
        
        '处理卫生材料系统的相关模块
        '处理卫材目录管理
        Call UpdateParameterValue(rsSys!系统, 1711, "私有模块\" & mstr用户名 & "\材料显示模式", "显示下级", "包含下级卫材")
        Call UpdateParameterValue(rsSys!系统, 1711, "公共模块\zl9Stuff\卫材增加模式", "品种", "品种增加模式")
        Call UpdateParameterValue(rsSys!系统, 1711, "公共模块\zl9Stuff\卫材增加模式", "品种->规格", "品种规格模式")
        Call UpdateParameterValue(rsSys!系统, 1711, "公共模块\zl9Stuff\卫材增加模式", "规格", "规格增加模式")
        Call UpdateParameterValue(rsSys!系统, 1711, "公共模块\zl9Stuff\卫生材料规格编辑", "指导差价率", "上次指导差价率")
        Call UpdateParameterValue(rsSys!系统, 1711, "公共模块\zl9Stuff\卫生材料规格编辑", "加成率", "上次加成率")
        '外购外购入库
        Call UpdateParameterValue(rsSys!系统, 1712, "私有模块\" & mstr用户名 & "\界面设置\卫材外购入库单\BillEdit", "mshBill宽度", "单据列宽")
        Call UpdateParameterValue(rsSys!系统, 1712, "私有模块\" & mstr用户名 & "\界面设置\卫材外购入库单\BillEdit", "mshBill名称", "单据列头文本")
        
        '处理发放管理
        Call UpdateParameterValue(rsSys!系统, 1723, "私有模块\" & mstr用户名 & "\zl9Stuff\未发料清单", "单据格式", "发料单据打印格式")
        Call UpdateParameterValue(rsSys!系统, 1723, "公共模块\zl9Stuff\卫材发放管理", "打印方式", "发料打印提醒方式")
        Call UpdateParameterValue(rsSys!系统, 1723, "公共模块\zl9Stuff\卫材发放管理", "业务类型", "查询业务类型")
        Call UpdateParameterValue(rsSys!系统, 1723, "公共模块\zl9Stuff\按单据进行退料", "单据类型", "最后退料单据类型")
        
        '-- 诊疗项目管理
        Call UpdateParameterValue(rsSys!系统, 1054, "公共模块\zl9CISBase\诊疗项目增加", "连续", "诊疗项目连续增加")
        Call UpdateParameterValue(rsSys!系统, 1054, "私有模块\" & mstr用户名 & "\zl9CISBase\frmClinicLists", "显示停用项目", "显示停用项目")
        Call UpdateParameterValue(rsSys!系统, 1054, "私有模块\" & mstr用户名 & "\zl9CISBase\frmClinicFind", "匹配方式", "匹配方式")
        Call UpdateParameterValue(rsSys!系统, 1054, "私有模块\" & mstr用户名 & "\zl9CISBase\frmClinicFind", "查找范围", "查找范围")
        Call UpdateParameterValue(rsSys!系统, 1054, "私有模块\" & mstr用户名 & "\zl9CISBase\frmClinicFind", "区分大小写", "区分大小写")
        Call UpdateParameterValue(rsSys!系统, 1054, "私有模块\" & mstr用户名 & "\zl9CISBase\frmClinicFind", "查找别名", "查找别名")
        
        '- 1059 检验项目管理
        Call UpdateParameterValue(rsSys!系统, 1059, "公共模块\zl9CISBase\frmLabItems", "列表范围", "列表范围")
        '-  1062 质控品管理
        Call UpdateParameterValue(rsSys!系统, 1062, "公共模块\zl9CISBase\frmMassResEdit", "隐藏中文名", "隐藏中文名")
        
        '1028 检验技师工作站
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "标本范围", "标本范围")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "待核收范围", "待核收范围")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "历次检验范围", "历次检验范围")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "标本序号生成规则", "标本序号生成规则")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "历次检验范围指定开始日期", "历次检验范围指定开始日期")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "自动刷新", "自动刷新")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "核收忽略时间", "核收忽略时间")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "核收显示收费", "核收显示收费")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "核收允许双向", "核收允许双向")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "急诊标本", "急诊标本")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "按仪器项目核收", "按仪器项目核收")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "历史病人识别", "历史病人识别")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "自适应显示结果", "自适应显示结果")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "按上次输入的标本号累加", "按上次输入的标本号累加")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "只在核收登记时显示登记窗口", "只在核收登记时显示登记窗口")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "登记时不需要输入项目", "登记时不需要输入项目")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "手工项目按项目累加标本号", "手工项目按项目累加标本号")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "仪器数据文件", "仪器数据文件")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "文件提取仪器", "文件提取仪器")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "文件提取范围", "文件提取范围")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "文件提取开始日期", "文件提取开始日期")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "文件提取结束日期", "文件提取结束日期")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMain", "清空接收日志", "清空接收日志")
        
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabFilter", "使用组合查询", "使用组合查询")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabFilter", "组合查询", "组合查询")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabFilter", "是否使用时间", "是否使用时间")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabMain", "缺省科室ID", "缺省科室ID")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabMain", "过滤仪器", "过滤仪器")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabMain", "仪器小组", "仪器小组")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabMain", "显示待核收", "显示待核收")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabMain", "隐藏检验图形", "隐藏检验图形")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabMain", "图像宽度", "图像宽度")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork", "显示检验备注", "显示检验备注")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork", "使用条码扫描", "使用条码扫描")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork", "连续输入", "连续输入")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabTrack", "隐藏中文名", "隐藏中文名")
        '--frmAddPatient
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmAddPatient", "选择科室", "frmAddPatient_选择科室")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmAddPatient", "选择仪器", "frmAddPatient_选择仪器")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmAddPatient", "选择类别", "frmAddPatient_选择类别")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmAddPatient", "开单科室", "frmAddPatient_开单科室")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmAddPatient", "开单医生", "frmAddPatient_开单医生")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmAddPatient", "执行科室", "frmAddPatient_执行科室")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmAddPatient", "检验仪器", "frmAddPatient_检验仪器")
        '--frmBatchAction
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmBatchAction", "不打印被合并标本", "frmBatchAction_不打印被合并标本")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmBatchAction", "同一个病人合并为一个报告单打印", "frmBatchAction_同一个病人合并为一个报告单打印")
        '--frmLabAuditingLand
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabAuditingLand", "时限", "frmLabAuditingLand_时限")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabAuditingLand", "时间", "frmLabAuditingLand_时间")
        '--frmLabBarCodeBatPrint
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabBarCodeBatPrint", "科室名称Id", "frmLabBarCodeBatPrint_科室名称Id")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabBarCodeBatPrint", "标本ID", "frmLabBarCodeBatPrint_标本ID")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabBarCodeBatPrint", "采集方法", "frmLabBarCodeBatPrint_采集方法")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabBarCodeBatPrint", "执行状态", "frmLabBarCodeBatPrint_执行状态")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "\zl9LisWork\frmLabBarCodeBatPrint", "是否标记为完成", "frmLabBarCodeBatPrint_是否标记为完成")
        '--frmLabMainFindRePort
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMainFindRePort", "使用时间范围", "frmLabMainFindRePort_使用时间范围")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLabMainFindRePort", "rptFind", "frmLabMainFindRePort_rptFind")
        '--frmLabMainSizer
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "zl9LisWork\检验中", "门诊病人", "检验中_门诊病人")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "zl9LisWork\检验中", "住院病人", "检验中_住院病人")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "zl9LisWork\检验中", "无主标本", "检验中_无主标本")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "zl9LisWork\检验中", "已审标本", "检验中_已审标本")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "zl9LisWork\检验中", "未审标本", "检验中_未审标本")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "zl9LisWork\检验中", "体检病人", "检验中_体检病人")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "zl9LisWork\检验中", "紧急医嘱", "检验中_紧急医嘱")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "zl9LisWork\检验中", "紧急标本", "检验中_紧急标本")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "zl9LisWork\待核收", "门诊病人", "待核收_门诊病人")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\" & mstr用户名 & "zl9LisWork\待核收", "住院病人", "待核收_住院病人")
        '--frmLabMB
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "仪器ID", "frmLabMB_仪器ID")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "阴性对照", "frmLabMB_阴性对照")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "通讯口", "frmLabMB_通讯口")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "波特率", "frmLabMB_波特率")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "数据位", "frmLabMB_数据位")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "停止位", "frmLabMB_停止位")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "校验位", "frmLabMB_校验位")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "波长", "frmLabMB_波长")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "参考波长", "frmLabMB_参考波长")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "振板频率", "frmLabMB_振板频率")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "振板时间", "frmLabMB_振板时间")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "进板方式", "frmLabMB_进板方式")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "空白形式", "frmLabMB_空白形式")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "项目ID", "frmLabMB_项目ID")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "减空白对照", "frmLabMB_减空白对照")
        Call UpdateParameterValue(rsSys!系统, 1208, "私有模块\zl9LisWork\frmLabMB", "阴性对照", "frmLabMB_阴性对照")
        '--frmLisStationWrite
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLisStationWrite", "查看原始结果", "frmLisStationWrite_查看原始结果")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLisStationWrite", "查看上次结果", "frmLisStationWrite_查看上次结果")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLisStationWrite", "查看标志", "frmLisStationWrite_查看标志")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLisStationWrite", "查看单位", "frmLisStationWrite_查看单位")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLisStationWrite", "查看参考", "frmLisStationWrite_查看参考")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLisStationWrite", "查看酶标", "frmLisStationWrite_查看酶标")
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLisStationWrite", "查看中文", "frmLisStationWrite_查看中文")
        '--frmLisStationWrite2
        Call UpdateParameterValue(rsSys!系统, 1208, "公共模块\zl9LisWork\frmLisStationWrite2", "查看上次结果", "frmLisStationWrite2_查看上次结果")
        
        '--1211
        Call UpdateParameterValue(rsSys!系统, 1211, "私有模块\" & mstr用户名 & "\frmLabSamplingFilter", "采集工作站过滤", "采集工作站过滤")
        Call UpdateParameterValue(rsSys!系统, 1211, "公共模块\zl9LisWork\frmLabSampling", "科室", "科室")
        Call UpdateParameterValue(rsSys!系统, 1211, "公共模块\zl9LisWork\frmLabSampling", "生成条码后打印", "生成条码后打印")
        Call UpdateParameterValue(rsSys!系统, 1211, "公共模块\zl9LisWork\frmLabSampling", "生成后标记为已完成", "生成后标记为已完成")
        Call UpdateParameterValue(rsSys!系统, 1211, "公共模块\zl9LisWork\frmLabSampling", "已完成后打印回执单", "已完成后打印回执单")
        Call UpdateParameterValue(rsSys!系统, 1211, "公共模块\zl9LisWork\frmLabSampling", "连续输入", "连续输入")
        Call UpdateParameterValue(rsSys!系统, 1211, "公共模块\zl9LisWork\frmLabSampling", "查找病人后光标移动", "查找病人后光标移动")
        Call UpdateParameterValue(rsSys!系统, 1211, "私有模块\" & mstr用户名 & "\frmLabSamplingRegister", "采集工作站登记", "采集工作站登记")
        '--1212
        Call UpdateParameterValue(rsSys!系统, 1211, "私有模块\" & mstr用户名 & "\frmLabSampleRegister", "是否按病区显示", "是否按病区显示")
        Call UpdateParameterValue(rsSys!系统, 1211, "私有模块\" & mstr用户名 & "\frmLabSampleRegisterFilter", "标本登记过滤", "标本登记过滤")
        Call UpdateParameterValue(rsSys!系统, 1211, "私有模块\" & mstr用户名 & "\frmLabSampleRegister", "科室", "科室")
        
        '-- 1209 历史质控查询
        Call UpdateParameterValue(rsSys!系统, 1209, "公共模块\zl9LisWork\frmQCHistory", "隐藏质控品", "隐藏质控品")
        Call UpdateParameterValue(rsSys!系统, 1209, "公共模块\zl9LisWork\frmQCHistory", "显示所有失控项目", "显示所有失控项目")
        Call UpdateParameterValue(rsSys!系统, 1209, "公共模块\zl9LisWork\frmQCHistory", "科室", "科室")
        Call UpdateParameterValue(rsSys!系统, 1209, "公共模块\zl9LisWork\frmQCHistory", "仪器", "仪器")
        Call UpdateParameterValue(rsSys!系统, 1209, "公共模块\zl9LisWork\frmQCHistory", "项目", "项目")
        
        '--病人入院管理
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "显示入住病人", "显示入住病人")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient\frmManageHosReg", "刷新方式", "刷新方式")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient\frmManageHosReg", "显示病人方式", "显示病人方式")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient\frmHosReg", "发卡模式", "发卡模式")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "当前预交票据号", "当前预交票据号")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "国籍", "国籍")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "民族", "民族")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "学历", "学历")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "婚姻状况", "婚姻状况")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "职业", "职业")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "身份", "身份")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "出生日期", "出生日期")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "身份证号", "身份证号")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "出生地点", "出生地点")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "家庭地址", "家庭地址")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "户口邮编", "户口邮编")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "家庭电话", "家庭电话")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "联系人姓名", "联系人姓名")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "联系人关系", "联系人关系")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "联系人地址", "联系人地址")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "联系人电话", "联系人电话")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "工作单位", "工作单位")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "单位电话", "单位电话")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "单位邮编", "单位邮编")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "单位开户行", "单位开户行")
        Call UpdateParameterValue(rsSys!系统, 1131, "私有模块\" & mstr用户名 & "\zl9InPatient", "单位帐号", "单位帐号")
        Call UpdateParameterValue(rsSys!系统, 1131, "公共模块\zl9InPatient", "共用就诊卡批次", "共用就诊卡批次")
        Call UpdateParameterValue(rsSys!系统, 1131, "公共模块\zl9InPatient", "共用预交票据批次", "共用预交票据批次")

        '--病人入出管理
        Call UpdateParameterValue(rsSys!系统, 1132, "私有模块\" & mstr用户名 & "\zl9InPatient", "当天入院", "当天入院")
        Call UpdateParameterValue(rsSys!系统, 1132, "公共模块\zl9InPatient", "待入科病人科室", "待入科病人科室")
        Call UpdateParameterValue(rsSys!系统, 1132, "公共模块\zl9InPatient", "入院天数", "入院天数")
        Call UpdateParameterValue(rsSys!系统, 1132, "公共模块\zl9InPatient", "出院天数", "出院天数")
        
        '--预交款管理
        Call UpdateParameterValue(rsSys!系统, 1103, "私有模块\" & mstr用户名 & "\zl9Patient", "当前预交票据号", "当前预交票据号")
        Call UpdateParameterValue(rsSys!系统, 1103, "公共模块\zl9Patient", "共用预交票据批次", "共用预交票据批次")
        Call UpdateParameterValue(rsSys!系统, 1103, "公共模块\zl9Patient", "LED显示欢迎信息", "LED显示欢迎信息")
        
        '--就诊卡管理
        Call UpdateParameterValue(rsSys!系统, 1102, "公共模块\zl9Patient", "共用就诊卡批次", "共用就诊卡批次")
        Call UpdateParameterValue(rsSys!系统, 1102, "公共模块\zl9Patient", "LED显示欢迎信息", "LED显示欢迎信息")
        
        '--病人信息管理
        Call UpdateParameterValue(rsSys!系统, 1101, "私有模块\" & mstr用户名 & "\zl9Patient", "当前预交票据号", "当前预交票据号")
        Call UpdateParameterValue(rsSys!系统, 1101, "私有模块\" & mstr用户名 & "\zl9Patient\frmManagePatient", "显示病人方式", "显示病人方式")
        Call UpdateParameterValue(rsSys!系统, 1101, "私有模块\" & mstr用户名 & "\zl9Patient\frmManagePatient", "病人类型", "病人类型")
        Call UpdateParameterValue(rsSys!系统, 1101, "私有模块\" & mstr用户名 & "\zl9Patient\frmPatient", "发卡模式", "发卡模式")
        Call UpdateParameterValue(rsSys!系统, 1101, "公共模块\zl9Patient", "共用会员卡批次", "共用会员卡批次")
        Call UpdateParameterValue(rsSys!系统, 1101, "公共模块\zl9Patient", "共用就诊卡批次", "共用就诊卡批次")
        Call UpdateParameterValue(rsSys!系统, 1101, "公共模块\zl9Patient", "共用预交票据批次", "共用预交票据批次")
        Call UpdateParameterValue(rsSys!系统, 1101, "公共模块\zl9Patient", "LED显示欢迎信息", "LED显示欢迎信息")
        
        '--合约单位管理1100
        Call UpdateParameterValue(rsSys!系统, 1100, "私有模块\" & mstr用户名 & "\zl9Patient\frmUnit\Menu", "mnuViewShowAll状态", "显示所有下级", True)
        Call UpdateParameterValue(rsSys!系统, 1100, "私有模块\" & mstr用户名 & "\zl10Patient\frmUnit\Menu", "mnuViewShowStop状态", "显示停用单位", True)

        '--电子病历管理
        Call UpdateParameterValue(rsSys!系统, 1070, "私有模块\" & mstr用户名 & "\zlRichEPR\自动缓存", "AutoSave", "AutoSave", True)
        Call UpdateParameterValue(rsSys!系统, 1070, "私有模块\" & mstr用户名 & "\zlRichEPR\自动缓存", "UndoLimit", "UndoLimit")
        Call UpdateParameterValue(rsSys!系统, 1070, "私有模块\" & mstr用户名 & "\zlRichEPR\自动缓存", "SaveInterval", "aveInterval")
        Call UpdateParameterValue(rsSys!系统, 1070, "私有模块\" & mstr用户名 & "\zlRichEPR\自动缓存", "AutoSaveEPR", "AutoSaveEPR", True)
        Call UpdateParameterValue(rsSys!系统, 1070, "私有模块\" & mstr用户名 & "\zlRichEPR\自动缓存", "SaveIntervalEPR", "SaveIntervalEPR")
        Call UpdateParameterValue(rsSys!系统, 1070, "私有模块\" & mstr用户名 & "\zlRichEPR\自动缓存", "AutoPageCount", "AutoPageCount", True)
        Call UpdateParameterValue(rsSys!系统, 1070, "私有模块\" & mstr用户名 & "\zlRichEPR\自动缓存", "AutoPageNote", "AutoPageNote", True)
        Call UpdateParameterValue(rsSys!系统, 1070, "私有模块\" & mstr用户名 & "\zlRichEPR\历史内容", "SharePageCount", "SharePageCount")
        Call UpdateParameterValue(rsSys!系统, 1070, "私有模块\" & mstr用户名 & "\zlRichEPR\静默打印", "NoAsk", "NoAsk", True)
        
        '--影像医技工作站
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frm3DSetup", "启用三维重建", "启用三维重建")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frm3DSetup", "3D程序路径", "3D程序路径")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frm3DSetup", "3D参数", "3D参数")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frm3DSetup", "3D功能", "3D功能")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "连续登记申请", "连续登记申请")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "登记直接检查", "登记直接检查")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "报到后自动打印申请单", "报到后自动打印申请单")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "不显示附加主述", "不显示附加主述")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "不显示造影剂 ", "不显示造影剂")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "开始检查自动打开报告", "开始检查自动打开报告")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "报告时观片", "报告时观片")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "不显示被取消的登记", "不显示被取消的登记")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "病人跟踪", "病人跟踪")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "只处理选中执行间病人", "只处理选中执行间病人")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "执行间范围", "执行间范围")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "病人列表表头字体", "病人列表表头字体")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "病人列表表头字号", "病人列表表头字号")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "病人列表表头粗体", "病人列表表头粗体")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "病人列表表头斜体", "病人列表表头斜体")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "病人列表内容字体", "病人列表内容字体")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "病人列表内容字号", "病人列表内容字号")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "病人列表内容粗体", "病人列表内容粗体")
        Call UpdateParameterValue(rsSys!系统, 1290, "公共模块\zl9PACSWork\frmPACSTechnicSetup", "病人列表内容斜体", "病人列表内容斜体")

        '--影像采集工作站
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "连续登记申请", "连续登记申请")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "登记直接检查", "登记直接检查")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "报到后自动打印申请单", "报到后自动打印申请单")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "不显示造影剂", "不显示造影剂")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "不显示附加主述", "不显示附加主述")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "开始检查自动打开报告", "开始检查自动打开报告")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "不显示被取消的登记", "不显示被取消的登记")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "报告时观片", "报告时观片")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "病人跟踪", "病人跟踪")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "只处理选中执行间病人", "只处理选中执行间病人")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "执行间范围", "执行间范围")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoCapture", "脚踏端口", "脚踏端口")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoCapture", "脚踏采集方式", "脚踏采集方式")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoCapture", "脚踏时间间隔", "脚踏时间间隔")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoCapture", "鼠标移动时显示大图", "鼠标移动时显示大图")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoCapture", "采集大图放大倍数", "采集大图放大倍数")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "病人列表表头字体", "病人列表表头字体")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "病人列表表头字号", "病人列表表头字号")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "病人列表表头粗体", "病人列表表头粗体")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "病人列表表头斜体", "病人列表表头斜体")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "病人列表内容字体", "病人列表内容字体")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "病人列表内容字号", "病人列表内容字号")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "病人列表内容粗体", "病人列表内容粗体")
        Call UpdateParameterValue(rsSys!系统, 1291, "公共模块\zl9PACSWork\frmVideoTechnicSetup", "病人列表内容斜体", "病人列表内容斜体")
        
        '--影像病理工作站
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "连续登记申请", "连续登记申请")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "登记直接检查", "登记直接检查")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "报到后自动打印申请单", "报到后自动打印申请单")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "不显示造影剂", "不显示造影剂")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "不显示附加主述", "不显示附加主述")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "开始检查自动打开报告", "开始检查自动打开报告")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "报告时观片", "报告时观片")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "不显示被取消的登记", "不显示被取消的登记")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "病人跟踪", "病人跟踪")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "只处理选中执行间病人", "只处理选中执行间病人")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "执行间范围", "执行间范围")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmVideoCapture", "脚踏端口", "脚踏端口")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmVideoCapture", "脚踏采集方式", "脚踏采集方式")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmVideoCapture", "脚踏时间间隔", "脚踏时间间隔")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmVideoCapture", "鼠标移动时显示大图", "鼠标移动时显示大图")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmVideoCapture", "采集大图放大倍数", "采集大图放大倍数")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "病人列表表头字体", "病人列表表头字体")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "病人列表表头字号", "病人列表表头字号")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "病人列表表头粗体", "病人列表表头粗体")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "病人列表表头斜体", "病人列表表头斜体")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "病人列表内容字体", "病人列表内容字体")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "病人列表内容字号", "病人列表内容字号")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "病人列表内容粗体", "病人列表内容粗体")
        Call UpdateParameterValue(rsSys!系统, 1293, "公共模块\zl9PACSWork\frmPathologyTechnicSetup", "病人列表内容斜体", "病人列表内容斜体")
        
        '基础部件药品目录管理
        Call UpdateParameterValue(rsSys!系统, 1023, "公共模块\zl9CisBase\药品增加模式", "品种增加模式", "品种增加模式")
        Call UpdateParameterValue(rsSys!系统, 1023, "公共模块\zl9CisBase\药品增加模式", "规格增加模式", "规格增加模式")
        
        '药房发药和药品流通部分
        '药品处方发药
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "收费处方显示方式", "收费处方显示方式")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "记帐处方显示方式", "记帐处方显示方式")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "查询天数", "查询天数")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "打印包含记帐单", "打印包含记帐单")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "打印退费单据间隔", "打印退费单据间隔")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "打印延迟", "打印延迟")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "刷新间隔", "刷新间隔")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "打印间隔", "打印间隔")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "显示付数", "显示付数")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "发药后自动打印", "发药后自动打印")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "药房属性", "药房属性")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "发现新单据是否打印", "发现新单据是否打印")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "打印指定发药窗口", "打印指定发药窗口")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "打印药品标签", "打印药品标签")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "发药窗口", "发药窗口")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "发药药房", "发药药房")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "配药人", "配药人")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "来源科室", "来源科室")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "自动配药", "自动配药")
        Call UpdateParameterValue(rsSys!系统, 1341, "公共模块\操作\zl9DrugStore\frm药品发药管理", "自动配药时限", "自动配药时限")
        Call UpdateParameterValue(rsSys!系统, 1341, "私有模块\" & mstr用户名 & "\zl9DrugStore\frm药品发药管理", "显示大小单位", "显示大小单位")
        
        '药品部门发药
        Call UpdateParameterValue(rsSys!系统, 1342, "公共模块\操作\zl9DrugStore\Frm部门发药管理", "发药药房", "发药药房")
        Call UpdateParameterValue(rsSys!系统, 1342, "公共模块\操作\zl9DrugStore\Frm部门发药管理", "自动打印", "自动打印")
        Call UpdateParameterValue(rsSys!系统, 1342, "私有模块\" & mstr用户名 & "\zl9DrugStore\部门发药管理", "显示大小单位", "显示大小单位")
        Call UpdateParameterValue(rsSys!系统, 1342, "私有模块\" & mstr用户名 & "\zl9DrugStore\部门发药管理", "按科室汇总显示汇总清单", "按科室汇总显示汇总清单")
        Call UpdateParameterValue(rsSys!系统, 1342, "私有模块\" & mstr用户名 & "\zl9DrugStore\部门发药管理", "操作模式", "操作模式")
        Call UpdateParameterValue(rsSys!系统, 1342, "私有模块\" & mstr用户名 & "\zl9DrugStore\部门发药管理", "记帐人", "记帐人")
        Call UpdateParameterValue(rsSys!系统, 1342, "私有模块\" & mstr用户名 & "\zl9DrugStore\部门发药管理", "毒理分类", "毒理分类")
        Call UpdateParameterValue(rsSys!系统, 1342, "私有模块\" & mstr用户名 & "\zl9DrugStore\部门发药管理", "价值分类", "价值分类")
        
        '药品库存查询
        Call UpdateParameterValue(rsSys!系统, 1309, "私有模块\" & mstr用户名 & "\zl9MediStore\药品库存查询", "单位", "单位")
        Call UpdateParameterValue(rsSys!系统, 1309, "私有模块\" & mstr用户名 & "\zl9MediStore\药品库存查询", "是否显示无库存药品", "是否显示无库存药品")
        Call UpdateParameterValue(rsSys!系统, 1309, "私有模块\" & mstr用户名 & "\zl9MediStore\药品库存查询", "效期报警月数", "效期报警月数")
        Call UpdateParameterValue(rsSys!系统, 1309, "私有模块\" & mstr用户名 & "\zl9MediStore\药品库存查询", "是否显示停用药品", "是否显示停用药品")
        
        '1560-病案审查
        Call UpdateParameterValue(rsSys!系统, 1560, "私有模块\" & mstr用户名 & "\zl9CISAudit\frmChildQuestion", "当前病人", "当前病人")
        Call UpdateParameterValue(rsSys!系统, 1560, "私有模块\" & mstr用户名 & "\zl9CISAudit\frmEPRAuditMan", "显示无业务科室", "显示无业务科室")
        
        '1561-病案借阅
        Call UpdateParameterValue(rsSys!系统, 1561, "私有模块\" & mstr用户名 & "\zl9CISAudit\frmSearchPatient", "常用条件", "常用条件")
        
        '1562-病案评分
        Call UpdateParameterValue(rsSys!系统, 1562, "私有模块\zl9CISAudit\frm病案评分", "未评分", "未评分")
        Call UpdateParameterValue(rsSys!系统, 1562, "私有模块\zl9CISAudit\frm病案评分", "未审核", "未审核")
        Call UpdateParameterValue(rsSys!系统, 1562, "私有模块\zl9CISAudit\frm病案评分", "已审核", "已审核")
        Call UpdateParameterValue(rsSys!系统, 1562, "私有模块\zl9CISAudit\frm病案评分", "定位范围", "定位范围")
        
    End If
    
    '处理物资系统的参数值升级
    '-----------------------------------------------------------------
    rsSys.Filter = "系统=4": rsUpgrade.Filter = "系统=4"
    If Not rsSys.EOF And rsUpgrade.EOF Then
        '处理物资系统的相关模块
        '处理物资目录管理
        Call UpdateParameterValue(rsSys!系统, 603, "私有模块\" & mstr用户名 & "\ZL9Material\界面\物资目录管理\卡片", "卡片", "显示隐藏卡片")
        Call UpdateParameterValue(rsSys!系统, 603, "私有模块\" & mstr用户名 & "\ZL9Material\界面\物资目录管理", "包含下级物资", "包含下级物资")
        Call UpdateParameterValue(rsSys!系统, 603, "私有模块\" & mstr用户名 & "\ZL9Material\界面\物资目录管理", "包含停用物资", "包含停用物资")
        
        '外购入库
        Call UpdateParameterValue(rsSys!系统, 309, "私有模块\" & mstr用户名 & "\ZL9Material\物资外购入库单\BillEdit", "mshBill宽度", "单据列宽")
        Call UpdateParameterValue(rsSys!系统, 309, "私有模块\" & mstr用户名 & "\ZL9Material\物资外购入库单\BillEdit", "mshBill名称", "单据列头文本")
        '领用管理
        Call UpdateParameterValue(rsSys!系统, 312, "私有全局\" & mstr用户名 & "\物资领用单\mshBill", "审核标志", "审核标志列宽")
    End If
    
    '处理设备系统的参数值升级
    '-----------------------------------------------------------------
    rsSys.Filter = "系统=6": rsUpgrade.Filter = "系统=6"
    If Not rsSys.EOF And rsUpgrade.EOF Then
        '处理设备系统的相关模块
        '设备目录管理
        Call UpdateParameterValue(rsSys!系统, 602, "私有模块\" & mstr用户名 & "\ZL9Device\设备目录管理\卡片", "显示卡片", "显示隐藏卡片")
        Call UpdateParameterValue(rsSys!系统, 602, "私有模块\" & mstr用户名 & "\ZL9Device\设备目录管理", "显示停用", "包含停用设备")
        
        '设备使用状态管理
        Call UpdateParameterValue(rsSys!系统, 603, "私有模块\" & mstr用户名 & "\ZL9Device\设备使用状态管理", "显示卡片", "显示隐藏卡片")
        '设备外购入库管理
        Call UpdateParameterValue(rsSys!系统, 616, "私有模块\" & mstr用户名 & "\ZL9Device\设备多选器-616", "设备选择器", "设备信息列表")
        Call UpdateParameterValue(rsSys!系统, 616, "私有模块\" & mstr用户名 & "\ZL9Device\设备多选器-616", "批次选择器", "批次信息列表")
        '设备调拨管理
        Call UpdateParameterValue(rsSys!系统, 618, "私有模块\" & mstr用户名 & "\ZL9Device\设备调拔单", "显示卡片", "显示卡片信息")
        '设备领用管理
        Call UpdateParameterValue(rsSys!系统, 619, "私有模块\" & mstr用户名 & "\ZL9Device\设备领用单", "显示卡片", "显示卡片信息")
        '设备使用管理
        Call UpdateParameterValue(rsSys!系统, 624, "私有模块\" & mstr用户名 & "\ZL9Device\设备使用管理", "包含附属设备", "包含附属设备")
        '设备保养管理
        Call UpdateParameterValue(rsSys!系统, 625, "私有模块\" & mstr用户名 & "\ZL9Device\设备保养管理", "包含附属设备", "包含附属设备")
        '设备检查管理
        Call UpdateParameterValue(rsSys!系统, 626, "私有模块\" & mstr用户名 & "\ZL9Device\设备检查管理", "包含附属设备", "包含附属设备")
        '设备维修管理
        Call UpdateParameterValue(rsSys!系统, 627, "私有模块\" & mstr用户名 & "\ZL9Device\设备维修管理", "包含附属设备", "包含附属设备")
        '设备变动管理
        Call UpdateParameterValue(rsSys!系统, 628, "私有模块\" & mstr用户名 & "\ZL9Device\设备变动管理", "包含附属设备", "包含附属设备")
        '设备下帐管理
        Call UpdateParameterValue(rsSys!系统, 631, "私有模块\" & mstr用户名 & "\ZL9Device\设备下帐管理", "包含附属设备", "包含附属设备")
        '在用设备查贸易
        Call UpdateParameterValue(rsSys!系统, 636, "私有模块\" & mstr用户名 & "\ZL9Device\设备在用查询", "分类ID", "上次设备分类ID")
        Call UpdateParameterValue(rsSys!系统, 636, "私有模块\" & mstr用户名 & "\ZL9Device\设备在用查询", "部门ID", "上次部门ID")
        Call UpdateParameterValue(rsSys!系统, 636, "私有模块\" & mstr用户名 & "\ZL9Device\设备在用查询", "包含下级部门", "包含下级部门")
        Call UpdateParameterValue(rsSys!系统, 636, "私有模块\" & mstr用户名 & "\ZL9Device\设备在用查询", "包含部门及设备", "同时限制部门及设备分类")
        
        '设备库存查询
        Call UpdateParameterValue(rsSys!系统, 637, "私有模块\" & mstr用户名 & "\ZL9Device\Frm设备库存查询\Menu", "显示库存物资", "仅显示库存设备")
        '设备明细帐
        Call UpdateParameterValue(rsSys!系统, 650, "私有模块\" & mstr用户名 & "\ZL9Device\设备明细帐", "只显示有数据的部门", "只显示有数据的部门")
    End If
    
    
    '处理病案系统的参数值升级
    '-----------------------------------------------------------------
    rsSys.Filter = "系统=23": rsUpgrade.Filter = "系统=23"
    If Not rsSys.EOF And rsUpgrade.EOF Then
        
    End If
    
    '处理院感系统的参数值升级
    '-----------------------------------------------------------------
    If Not rsSys.EOF And rsUpgrade.EOF Then
        Call UpdateParameterValue(rsSys!系统, 2301, "私有模块\" & mstr用户名 & "\ZL9Device\zl9Infect\污水项目管理", "显示停用", "污水项目-显示停用")
        Call UpdateParameterValue(rsSys!系统, 2301, "私有模块\" & mstr用户名 & "\ZL9Device\zl9Infect\污水项目管理", "显示下级", "污水项目-显示下级")
        Call UpdateParameterValue(rsSys!系统, 2301, "私有模块\" & mstr用户名 & "\ZL9Device\zl9Infect\消毒项目管理", "显示停用", "消毒项目-显示停用")
        Call UpdateParameterValue(rsSys!系统, 2301, "私有模块\" & mstr用户名 & "\ZL9Device\zl9Infect\消毒项目管理", "显示下级", "消毒项目-显示下级")
        Call UpdateParameterValue(rsSys!系统, 2301, "私有模块\" & mstr用户名 & "\ZL9Device\zl9Infect\易感因素管理", "显示停用", "易感因素-显示停用")
        Call UpdateParameterValue(rsSys!系统, 2301, "私有模块\" & mstr用户名 & "\ZL9Device\zl9Infect\易感因素管理", "显示下级", "易感因素-显示下级")
    End If
    Exit Sub
errH:
    If zlComLib.ErrCenter() = 1 Then Resume
    Call zlComLib.SaveErrLog
End Sub

Private Sub UpdateParameterValue(ByVal int系统 As Integer, ByVal int模块 As Integer, _
    ByVal strPath As String, ByVal strPreName As String, ByVal strNowName As String, Optional ByVal blnTransBool As Boolean)
'功能：更新具体某个注册表参数的值
'参数：int系统=系统号,因为以前注册表参数是不分帐套存储的，因此升级时只处理标准的系统号
'      int模块=模块号
'      strPath=原注册表存放路径
'      strPreName=原注册表存放的参数名(注册表键名)
'      strNowName=新的存放到数据库中的本机参数名
'      blnTransBool=是否将注册表存放为True或False的值转为1或0存到数据库中
    Dim strVal As String, strSQL As String
    
    On Error GoTo errH
    
    strVal = GetSetting("ZLSOFT", strPath, strPreName)
    If strVal = "" Then Exit Sub '没有值时也当成无此键处理。
    If blnTransBool Then
        If UCase(strVal) = "TRUE" Then
            strVal = "1"
        ElseIf UCase(strVal) = "FALSE" Then
            strVal = "0"
        End If
    End If
    
    strSQL = "zl_Parameters_Update('" & strNowName & "','" & Replace(strVal, "'", "''") & "'," & int系统 * 100 & "," & int模块 & ")"
    zlDataBase.ExecuteProcedure strSQL, "UpdateParameterValue"
    
    DeleteSetting "ZLSOFT", strPath, strPreName '不存在时删除会出错
    Exit Sub
errH:
    If zlComLib.ErrCenter = 1 Then Resume
    Call zlComLib.SaveErrLog
End Sub
