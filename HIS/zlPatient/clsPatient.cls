VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit '要求变量声明
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前数据库登录用户名

    Set gcnOracle = cnMain
    Set gfrmMain = frmMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrUnitName = GetUnitName
    
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not CheckValid Then Exit Sub

    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul

    If glngSys Like "8??" Then App.Title = "客户关系管理"
    
    '-------------------------------------------------
    '1100       合约单位管理
    '1101       病人信息管理
    '1102       就诊卡发放管理
    '1103       预交款管理
    '1104       预交款操作日报
    '1105       合约单位费用
    '1106       病人费用审批
    Call InitLocPar(lngModul)
    Call CreateSquareCardObject(frmMain, lngModul)
    Select Case lngModul
        Case P合约单位管理 '合约单位管理
            Call InitLocPar(lngModul)
            frmUnit.Show , frmMain
        Case P病人信息管理 '病人信息管理
            Call InitSysPar
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            frmManagePatient.Show , frmMain
'        Case P就诊卡发放管理 '就诊卡发放管理
'            Call InitSysPar
'            frmManageIDCard.Show , frmMain
        Case P预交款管理  '预交款管理
            Call InitSysPar
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            frmManageDeposit.Show , frmMain
        Case P病人费用审批  '病人费用审批
            Call InitLocPar(lngModul)
            Call InitSysPar
            frmManageExamine.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前数据库登录用户名

    Set gcnOracle = cnMain
    Set gfrmMain = Nothing
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrUnitName = GetUnitName
    
    gstrDBUser = strDBUser
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not CheckValid Then Exit Sub

    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
   
    If glngSys Like "8??" Then App.Title = "客户关系管理"
    
    '-------------------------------------------------
    '1100       合约单位管理
    '1101       病人信息管理
    '1102       就诊卡发放管理
    '1103       预交款管理
    '1104       预交款操作日报
    '1105       合约单位费用
    '1106       病人费用审批
    Call InitLocPar(lngModul)
    
    Select Case lngModul
        Case P合约单位管理 '合约单位管理
            Call InitLocPar(lngModul)
            Call CreateSquareCardObject(frmUnit, lngModul)
            zlCommFun.ShowChildWindow frmUnit.hWnd, lngMain
        Case P病人信息管理 '病人信息管理
            Call InitSysPar
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(frmManagePatient, lngModul)
            zlCommFun.ShowChildWindow frmManagePatient.hWnd, lngMain
'        Case P就诊卡发放管理 '就诊卡发放管理
'            Call InitSysPar
'            zlCommFun.ShowChildWindow frmManageIDCard.hwnd, lngMain
        Case P预交款管理 '预交款管理
            Call InitSysPar
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(frmManageDeposit, lngModul)
            zlCommFun.ShowChildWindow frmManageDeposit.hWnd, lngMain
        Case P病人费用审批 '病人费用审批
            Call InitLocPar(lngModul)
            Call InitSysPar
            Call CreateSquareCardObject(frmManageExamine, lngModul)
            zlCommFun.ShowChildWindow frmManageExamine.hWnd, lngMain
    End Select
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
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Function NurseDeposit(ByVal lngSys As Long, cnMain As ADODB.Connection, frmMain As Object, _
    ByVal strDBUser As String, Optional lng病人ID As Long, Optional lng主页ID As Long, _
    Optional ByVal bln余额退款 As Boolean = True, Optional ByVal bytPrepayType As Byte = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能： 护士站调用余额退款窗口
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '   strDBUser:当前数据库登录用户名
    '   lng病人ID-缺省的病人ID
    '   lng主页ID-缺省的主页ID
    '   bytPrepayType-预交类型(0-门诊和住院;1-门诊;2-住院)
    '返回:
    '编制:刘尔旋
    '日期:2016-09-05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNursePrivs As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
    glngSys = lngSys
    gstrPrivs = ";基本;住院预交;预交结清退款;预交退款;预交收款;负数缴款"
    If bytPrepayType <> 2 Then gstrPrivs = gstrPrivs & ";门诊预交;"
    strNursePrivs = GetPrivFunc(lngSys, 9000)
    
    Call GetUserInfo
    
    '-------------------------------------------------
    '1103      预交款收款窗口
    Call InitLocPar(1103)
    Call InitSysPar
    
    '医保接口
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(frmMain, 1103)
    
    If strNursePrivs = "" Then
        MsgBox "你不具备护士站使用余额退款的权限,请联系系统管理员!", vbInformation, gstrSysName
        Exit Function
    Else
        strSQL = "Select 1 From 病案主页 Where 病人ID=[1] And 主页ID=[2] And 出院日期 Is Null"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "在院状态", lng病人ID, lng主页ID)
        If rsTmp.EOF Then
            '出院
            If InStr(";" & strNursePrivs & ";", ";出院病人余额退款;") = 0 Then
                MsgBox "你不具备对出院病人使用余额退款的权限,请联系系统管理员!", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            '在院
            If InStr(";" & strNursePrivs & ";", ";在院病人余额退款;") = 0 Then
                MsgBox "你不具备对在院病人使用余额退款的权限,请联系系统管理员!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If

    NurseDeposit = frmDeposit.zlShowEdit(frmMain, 0, IIf(bln余额退款, 3, 0), gstrPrivs, 1103, _
        bytPrepayType, , , , lng病人ID, lng主页ID, 0, True)
End Function

Public Function PlusDeposit(ByVal lngSys As Long, cnMain As ADODB.Connection, frmMain As Object, _
    ByVal strDBUser As String, Optional bytCallObject As Byte = 0, _
    Optional lng病人ID As Long, Optional lng主页ID As Long, _
    Optional dblDefPrePayMoney As Double = 0, _
    Optional byt预交类型 As Integer = 0, _
    Optional blnOneCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能： 调用预交款收款窗口
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '   strDBUser:当前数据库登录用户名
    '  bytCallObject:刘兴洪加入(0-预交款调用(缺省的);1-病人费用查询调用,2-医疗卡调用)
    '  lng病人ID-缺省的病人ID
    '  lng主页ID-缺省的主页ID
    '  dblDefPrePayMoney-缺省的预付金额
    '  byt预交类型: 0-收预交款(缺省,可切换到退),1-浏览单据(1),2-作废状态(1); 3-余额退款(37770), 4-门诊转住院;5-住院转门诊
    '返回:
    '编制:刘兴洪
    '日期:2011-02-17 17:18:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt类型 As Byte
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, P预交款管理) & ";"
    Call GetUserInfo
    If gstrPrivs = ";;" Then
         MsgBox "你不具备缴预交款的权限,不能进行缴预交!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    
    '-------------------------------------------------
    '1103      预交款收款窗口
    Call InitLocPar(P预交款管理)
    Call InitSysPar
    
    '医保接口
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(frmMain, P预交款管理)
    If byt预交类型 = 4 Then
        byt类型 = 1
    ElseIf byt预交类型 = 5 Then
        byt类型 = 2: byt预交类型 = 4
    Else
        byt类型 = IIf(lng主页ID <> 0, 2, 1)
    End If
    '0-收预交款(缺省,可切换到退),1-浏览单据(1),2-作废状态(1);3-余额退款(37770)
    PlusDeposit = frmDeposit.zlShowEdit(frmMain, bytCallObject, byt预交类型, gstrPrivs, P预交款管理, byt类型, , , , lng病人ID, lng主页ID, dblDefPrePayMoney, , blnOneCard)
End Function

Public Function RefundDeposit(ByVal lngSys As Long, cnMain As ADODB.Connection, frmMain As Object, _
    ByVal strDBUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能： 调用余额退款窗口
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '   strDBUser:当前数据库登录用户名
    '返回:
    '编制:刘兴洪
    '日期:2011-02-17 17:18:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt类型 As Byte
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, P预交款管理) & ";"
    Call GetUserInfo
    If gstrPrivs = ";;" Then
         MsgBox "你不具备预交款管理的权限,不能进行余额退款!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    
    '-------------------------------------------------
    '1103      预交款收款窗口
    Call InitLocPar(P预交款管理)
    Call InitSysPar
    
    '医保接口
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(frmMain, P预交款管理)

    RefundDeposit = frmDeposit.zlShowEdit(frmMain, 0, 3, gstrPrivs, P预交款管理)
End Function

Public Function IDCard(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, _
    ByVal strDBUser As String, Optional lngModule As Long = 0, _
    Optional lng卡类别ID As Long, Optional lng病人ID As Long = 0, Optional strPrivs As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:直接调用发卡功能
    '返回:发卡成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-26 12:17:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardObject As Object, strExpend As String
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    glngSys = lngSys: glngModul = lngModule: gstrPrivs = strPrivs
    gstrDBUser = strDBUser: Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrUnitName = GetUnitName
    Set objCardObject = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        MsgBox "医疗卡部件不存在,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If objCardObject.zlInitComponents(frmMain, lngModule, lngSys, strDBUser, cnMain, False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Function
    End If
    '调用发卡
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        lngCardTypeID-医疗卡类别
    '        lng病人ID-传入指定的病人ID,对指定病人进行发卡
    '        strPrivs-可以传入空，传入空时取医疗卡管理模块的相关权限
    '问题:48416
    IDCard = objCardObject.zlSendCard(frmMain, lngModule, lng卡类别ID, lng病人ID, strPrivs)
    Set objCardObject = Nothing
End Function

Private Sub Class_Terminate()
    '62646:刘鹏飞,2013-06-18
    If Forms.Count = 0 Then
        Call CloseSquareCardObject '关闭卡结算信息
    End If
End Sub
