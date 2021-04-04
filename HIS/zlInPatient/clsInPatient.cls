VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInPatient"
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
'   strDBUser:当前登录数据库用户名
    Set gfrmMain = frmMain
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call InitLocPar(glngModul)
    Call CreateSquareCardObject(frmMain, lngModul)
    Select Case lngModul
        Case P病区床位管理 '病区床位管理
            Call InitSysPar
            frmManageBeds.Show , frmMain
        Case P病人入院管理 '病人入院管理
            Call InitSysPar
            
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            
            If gfrmMain Is Nothing Then
                frmManageHosReg.Show
            Else
                frmManageHosReg.Show , frmMain
            End If
        Case P病人入出管理 '病人入出管理
            Call InitSysPar
            
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)

            If gfrmMain Is Nothing Then
                frmManageCourse.Show
            Else
                frmManageCourse.Show , frmMain
            End If
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名
    Set gfrmMain = Nothing
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call InitLocPar(glngModul)
    Select Case lngModul
        Case P病区床位管理 '病区床位管理
            Call InitSysPar
            Call CreateSquareCardObject(frmManageBeds, lngModul)
            ZLCommFun.ShowChildWindow frmManageBeds.hWnd, lngMain
        Case P病人入院管理 '病人入院管理
            Call InitSysPar
            
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(frmManageHosReg, lngModul)
            ZLCommFun.ShowChildWindow frmManageHosReg.hWnd, lngMain
        Case P病人入出管理 '病人入出管理
            Call InitSysPar
            
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(frmManageCourse, lngModul)
            ZLCommFun.ShowChildWindow frmManageCourse.hWnd, lngMain
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

Public Sub InitCallByNurse(frmMain As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, lngSys As Long)
'功能：由护士工作站调用时初始化数据
    
    Set gfrmMain = frmMain
    'If Not CheckValid Then Exit Sub    '由于护士站中同时调用了其它部件,在调其它部件时把通讯原子删除了,所以这里不能再检查
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = P病人入出管理
    
    gstrPrivs = GetPrivFunc(glngSys, glngModul)
    gstrUnitName = GetUnitName
    
    Call GetUserInfo
    Call InitSysPar
    Call InitLocPar(glngModul)
    
    '医保接口
    Call gclsInsure.InitOracle(gcnOracle)
End Sub
Public Function zl_ExecPatiChange(ByVal bytFun As Byte, ByRef frmParent As Object, ByRef strPrivs As String, ParamArray arrPar() As Variant) As Boolean
'功能:执行病人变动相关功能
'参数:bytFun:0-入科
'     arrPar:根据不同的功能调用，传入不同的参数,目前最多支持10个参数
    Select Case UBound(arrPar) + 1
    Case 1
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0))
    Case 2
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1))
    Case 3
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2))
    Case 4
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3))
    Case 5
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4))
    Case 6
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4), arrPar(5))
    Case 7
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4), arrPar(5), arrPar(6))
    Case 8
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4), arrPar(5), arrPar(6), arrPar(7))
    Case 9
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4), arrPar(5), arrPar(6), arrPar(7), arrPar(8))
    Case 10
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4), arrPar(5), arrPar(6), arrPar(7), arrPar(8), arrPar(9))
    End Select
End Function
Public Sub zlOutFeeToInFee(ByVal frmMain As Object, ByVal cnMain As ADODB.Connection, _
    ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, strDBUser As String, _
    ByVal lng病人ID As Long, intPatientRange As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊费用转住院费用(门诊费用只能转到当前在院病人身上)
    '入参:frmMain-调用的主窗体
    '       strPrivs-权限串
    '       lngModule-模块号
    '       intPatientRange-(0-所有病人,1-任何费用未结清病人;2-体检未结清的病人;3-住院未结清的病人;4-门诊未结清的病人)
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-11-10 11:09:57
    '问题:33635
    '调用模块:
    '     住院结帐管理-->门诊费用转住院
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set gfrmMain = frmMain
    If UserInfo.ID = 0 Or glngSys = 0 Or gcnOracle Is Nothing Then
        Set gcnOracle = cnMain
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrDBUser = strDBUser
        glngSys = lngSys: glngModul = lngModule
        gstrPrivs = GetPrivFunc(glngSys, glngModul)
        gstrUnitName = GetUnitName
        Call GetUserInfo
        Call InitSysPar
        Call InitLocPar(glngModul)
        '医保接口
        Call gclsInsure.InitOracle(gcnOracle)
    End If
    If gobjSquare Is Nothing Then Call CreateSquareCardObject(frmMain, lngModule) '只调用一次,第二不进行调用
    Call frmChargeTurn.ShowMe(frmMain, lng病人ID, "", True, intPatientRange, strPrivs, lngModule)
End Sub


Public Sub zlPatiMemoReadAndWrite(ByVal frmParent As Object, ByVal cnMain As ADODB.Connection, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String, Optional ByVal blnEdit As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人备注信息编辑/查看
    '入参:frmParent-调用的主窗体
    '出参:
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If UserInfo.ID = 0 Or glngSys = 0 Or gcnOracle Is Nothing Then
        Set gcnOracle = cnMain
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrUnitName = GetUnitName
        Call GetUserInfo
    End If
    Call ExecPatiChange(EFun.E病人备注编辑, frmParent, strPrivs, lng病人ID, lng主页ID)
End Sub

Private Sub Class_Terminate()
    '62646:刘鹏飞,2013-06-18
    If Forms.Count = 0 Then
        CloseSquareCardObject '关闭卡结算部件
    End If
End Sub
