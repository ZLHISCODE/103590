VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If

    Set gfrmMain = frmMain
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    glngSys = lngSys
    gstrDBOwner = Sys.SystemOwner(lngSys)  ' 取系统所有者
    gstrUnitName = GetUnitName
    gstrMatch = IIf(Val(zlDatabase.GetPara("输入匹配", , , True)) = 0, "%", "")

    '取药品价格等级
    Call GetPriceClass
    
    '-------------------------------------------------
    On Error Resume Next
    Select Case lngModul
        Case 1023 '药品目录管理
            frmMediLists.ShowMe gfrmMain, False
        Case 1050 '诊疗参考管理
            frmCureRefers.Show , gfrmMain
        Case 1051 '疾病诊断参考
            frmDiagnoses.Show , gfrmMain
        Case 1053 '诊治所见项目
            frmVItemLists.Show , gfrmMain
        Case 1054 '诊疗项目管理
            frmClinicLists.Show , gfrmMain
        Case 1055 '医嘱频率设置
            frmAdviceFreq.Show , gfrmMain
        Case 1056 '影像检查项目
            frmRadLists.Show , gfrmMain
        Case 1057 '影像设备目录
            frmPacsDev.Show , gfrmMain
        Case 1049 '检查部位设置
            frmCheckPartList.Show , gfrmMain
        Case 1058   '检验仪器管理
            frmApparatusList.Show , gfrmMain
'            frmInstrument.Show , gfrmMain
        Case 1059   '检验项目管理
            frmLabItems.Show , gfrmMain
'            frmVerify.Show , gfrmMain
        Case 1060 '抗生素分组管理
            frmAntiList.Show , gfrmMain
'            FrmAntibioticManager.Show , gfrmMain
        Case 1061 '检验细菌管理
            FrmMicrobeList.Show , gfrmMain
'            FrmBacilliManager.Show , gfrmMain
        Case 1062 '质控品管理
            frmMassResList.Show , gfrmMain
'            FrmQualityResManager.Show , gfrmMain
        Case 1063 '质控规则管理
            frmMassRuleList.Show , gfrmMain
'            FrmQualityRuleManager.Show , gfrmMain
        Case 1065 '检验报告模板
            frmLabRptList.Show , gfrmMain
        Case 1066   '仪器质控规则
            frmAppRuleList.Show , gfrmMain
        Case 1067 '采血管管理 070425 add
            frmMiningVessels.Show , gfrmMain
        Case 1068 '审核规则 070928 add
            frmLabVerifyList.Show , gfrmMain
        Case 1069 '检验小组 090104
            frmLabGroup.Show , gfrmMain
        Case 1077 '基础信息管理
            frmBaseInfoList.Show , gfrmMain
        Case 1024 '抗菌药物授权
            frmKSSManagement.Show , gfrmMain
        Case 1276 '标准路径参考
            frmStPathMain.Show , gfrmMain
        Case 1080 '手术授权管理
            frmOPSEmpower.Show , gfrmMain
        Case 1081 '库房货位
            frmStoreSpace.Show , gfrmMain
        Case 1082 '医生授权管理
            frmDoctorManage.Show , gfrmMain
        Case 1085 '医生交接班内容设置
            frmDocShiftBase.Show , gfrmMain
    End Select
End Sub

Public Sub SetMedList(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, ByVal strDBUser As String)
'功能：科室设置临床自管药

    Set gfrmMain = frmMain
'    If Not CheckValid Then Exit Sub    '检查通讯原子
    
    Set gcnOracle = cnMain
    glngSys = lngSys
    glngModul = 1023
    gstrDBUser = strDBUser
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrPrivs = GetPrivFunc(lngSys, 1023)
    
    If gfrmMain Is Nothing Then
        frmMediLists.Show vbModal
    Else
        frmMediLists.ShowMe frmMain, True
    End If
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If

    Set gfrmMain = Nothing
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    glngSys = lngSys
    gstrDBOwner = Sys.SystemOwner(lngSys) ' 取系统所有者
    gstrUnitName = GetUnitName
    gstrMatch = IIf(Val(zlDatabase.GetPara("输入匹配", , , True)) = 0, "%", "")
    
    '取药品价格等级
    Call GetPriceClass
    
    '-------------------------------------------------
    On Error Resume Next
    Select Case lngModul
        Case 1023 '药品目录管理
            OS.ShowChildWindow frmMediLists.hwnd, lngMain
        Case 1050 '诊疗参考管理
            OS.ShowChildWindow frmCureRefers.hwnd, lngMain
        Case 1051 '疾病诊断参考
            OS.ShowChildWindow frmDiagnoses.hwnd, lngMain
        Case 1053 '诊治所见项目
            OS.ShowChildWindow frmVItemLists.hwnd, lngMain
        Case 1054 '诊疗项目管理
            OS.ShowChildWindow frmClinicLists.hwnd, lngMain
        Case 1055 '医嘱频率设置
            OS.ShowChildWindow frmAdviceFreq.hwnd, lngMain
        Case 1056 '影像检查项目
            OS.ShowChildWindow frmRadLists.hwnd, lngMain
        Case 1057 '影像设备目录
            OS.ShowChildWindow frmPacsDev.hwnd, lngMain
        Case 1049 '检查部位设置
            OS.ShowChildWindow frmCheckPartList.hwnd, lngMain
        Case 1058   '检验仪器管理
            OS.ShowChildWindow frmApparatusList.hwnd, lngMain
        Case 1059   '检验项目管理
            OS.ShowChildWindow frmLabItems.hwnd, lngMain
        Case 1060 '抗生素分组管理
            OS.ShowChildWindow frmAntiList.hwnd, lngMain
        Case 1061 '检验细菌管理
            OS.ShowChildWindow FrmMicrobeList.hwnd, lngMain
        Case 1062 '质控品管理
            OS.ShowChildWindow frmMassResList.hwnd, lngMain
        Case 1063 '质控规则管理
            OS.ShowChildWindow frmMassRuleList.hwnd, lngMain
        Case 1065 '检验报告模板
            OS.ShowChildWindow frmLabRptList.hwnd, lngMain
        Case 1066   '仪器质控规则
            OS.ShowChildWindow frmAppRuleList.hwnd, lngMain
        Case 1067 '采血管管理 070425 add
            OS.ShowChildWindow frmMiningVessels.hwnd, lngMain
        Case 1068 '审核规则 070928 add
            OS.ShowChildWindow frmLabVerifyList.hwnd, lngMain
        Case 1069 '检验小组 090104
            OS.ShowChildWindow frmLabGroup.hwnd, lngMain
        Case 1077 '基础信息管理
            OS.ShowChildWindow frmBaseInfoList.hwnd, lngMain
        Case 1024 '抗菌药物授权
            OS.ShowChildWindow frmKSSManagement.hwnd, lngMain
        Case 1276 '标准路径参考
            OS.ShowChildWindow frmStPathMain.hwnd, lngMain
        Case 1080 '手术授权管理
            OS.ShowChildWindow frmOPSEmpower.hwnd, lngMain
        Case 1081 '库房货位
            frmStoreSpace.Show , gfrmMain
        Case 1082 '医生授权管理
            OS.ShowChildWindow frmDoctorManage.hwnd, lngMain
        Case 1085 '医生交接班内容设置
            frmDocShiftBase.Show , gfrmMain
    End Select
End Sub

Public Sub CallClinicScheme(frmMain As Object, cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strDBUser As String, ByVal int范围 As Integer)
'功能：调用诊疗项目管理-成套方案管理
'参数：int范围=成套方案的可使用场合，1-门诊,2-住院,3-门诊和住院

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = 1054 '1054-诊疗项目管理
    
    '权限判断和处理
    gstrPrivs = ";" & GetPrivFunc(lngSys, 1054) & ";"
    If InStr(gstrPrivs, ";成套方案编辑;") = 0 Or InStr(gstrPrivs, ";管理成套方案;") = 0 Then
        MsgBox "你没有成套方案的管理或编辑权限，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrPrivs = Replace(gstrPrivs, ";项目导入;", ";")
    gstrPrivs = Replace(gstrPrivs, ";收费设置;", ";")
    gstrPrivs = Replace(gstrPrivs, ";管理诊疗项目;", ";")
    gstrPrivs = Replace(gstrPrivs, ";管理中药配方;", ";")
    
    Call GetUserInfo
    
    '初始化核心部件
    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs)
    
    frmClinicLists.ShowMeWithScheme frmMain, int范围
End Sub

Public Sub CallOPSEmpower(frmMain As Object, cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strDBUser As String)
'功能：医生站调用手术审核管理

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = 1054 '1054-诊疗项目管理
    
    '权限判断和处理
    gstrPrivs = GetPrivFunc(lngSys, 1080)
    If gstrPrivs = "" Then
        MsgBox "你没有手术授权管理权限，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Sub
    End If

    Call GetUserInfo
    
    frmOPSEmpower.Show , frmMain
End Sub

Public Sub CallSetClinicCharge(ByVal lngDeptID As Long, ByVal lngMode As Long, frmMain As Object, cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strDBUser As String, ByVal int场合 As Integer, ByVal blnModify As Boolean)
'功能：调用诊疗项目 病区 收费对照
'参数： lngDeptID 病区ID
'       lngMode=是否模态显示窗体
'       int场合= 1-门诊 ，2-住院
    '------------------------------------------------
    Dim lngModul As Long
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    
    If InStr(frmMain.Caption, "门诊医生工作站 -") > 0 Then
        lngModul = 1260
    ElseIf InStr(frmMain.Caption, "住院护士工作站 -") > 0 Then
        lngModul = 1262
    ElseIf InStr(frmMain.Caption, "医技工作站 -") > 0 Then
        lngModul = 1263
    ElseIf InStr(frmMain.Caption, "新版住院护士工作站") > 0 Then
        lngModul = 1265
    Else
        MsgBox "该模块不能调用“诊疗项目费用设置”功能！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul

    gstrDBOwner = Sys.SystemOwner(lngSys) ' 取系统所有者
    gstrUnitName = GetUnitName
    gstrMatch = IIf(Val(zlDatabase.GetPara("输入匹配", , , True)) = 0, "%", "")

    '权限判断和处理
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngMode) & ";"
'    If InStr(gstrPrivs, ";病区收费项目对照;") = 0 Then
'        MsgBox "您没有病区收费项目对照权限，请与系统管理员联系。", vbInformation, gstrSysName
'        Exit Sub
'    End If
    
    Call GetUserInfo
    Call frmSetClinicCharge.ShowMe(lngDeptID, lngMode, frmMain, int场合, Not blnModify)
End Sub


Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = gobjKernel.CloseWindows() And (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub


