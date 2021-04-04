VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit '要求变量声明
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名
'返回：
    Set gfrmMain = frmMain
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    gstrDBUser = strDbUser
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If

    '-------------------------------------------------
    '1110       挂号安排管理
    '1111       门诊挂号管理
    '1112       挂号员日报
    '1114       临床出诊安排
    Call InitLocPar(glngModul)
    Select Case glngModul
        Case 1110                 '挂号安排
            Call InitSysPar
            frmRegistPlan.Show , frmMain
        Case 1111                 '门诊挂号管理
            Call InitSysPar
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(gfrmMain, lngModul)
            Call CreatePublicExpenseObject(lngModul)
            If gfrmMain Is Nothing Then
                If gbytRegistMode = 0 Then
                    frmRegist.Show
                Else
                    frmRegistNew.Show
                End If
            Else
                If gbytRegistMode = 0 Then
                    frmRegist.Show , frmMain
                Else
                    frmRegistNew.Show , frmMain
                End If
            End If
        Case 1113
            Call InitSysPar
            Call CreateSquareCardObject(gfrmMain, lngModul)
            frmDistRoomManager.Show , frmMain     '门诊分诊管理
        Case 1114
            If gfrmMain Is Nothing Then
                frmClinicPlanMainManage.Show
            Else
                frmClinicPlanMainManage.Show , frmMain
            End If
        Case 1115
            Call InitSysPar
            Call CreateSquareCardObject(gfrmMain, lngModul)
            Call CreatePublicExpenseObject(lngModul)
            If gfrmMain Is Nothing Then
                frmServiceCenter.Show
            Else
                frmServiceCenter.Show , frmMain
            End If
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名
'返回：
    Set gfrmMain = Nothing
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    gstrDBUser = strDbUser
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If

    '-------------------------------------------------
    '1110       挂号安排管理
    '1111       门诊挂号管理
    '1112       挂号员日报
    '1114       临床出诊安排
    Call InitLocPar(glngModul)
    Select Case glngModul
        Case 1110                 '挂号安排
            Call InitSysPar
            OS.ShowChildWindow frmRegistPlan.Hwnd, lngMain
        Case 1111                 '门诊挂号管理
            Call InitSysPar
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(gfrmMain, lngModul)
            Call CreatePublicExpenseObject(lngModul)
            If gbytRegistMode = 0 Then
                OS.ShowChildWindow frmRegist.Hwnd, lngMain
            Else
                OS.ShowChildWindow frmRegistNew.Hwnd, lngMain
            End If
        Case 1113                 '门诊分诊管理
            Call InitSysPar
            Call CreateSquareCardObject(gfrmMain, lngModul)
            OS.ShowChildWindow frmDistRoomManager.Hwnd, lngMain
        Case 1114
            OS.ShowChildWindow frmClinicPlanMainManage.Hwnd, lngMain
        Case 1115
            Call InitSysPar
            Call CreateSquareCardObject(gfrmMain, lngModul)
            Call CreatePublicExpenseObject(lngModul)
            OS.ShowChildWindow frmServiceCenter.Hwnd, lngMain
    End Select
End Sub

Public Function PrintRegistBill(frmMain As Object, cnMain As ADODB.Connection, _
    ByVal lngSys As Long, ByVal strDbUser As String, _
    ByVal strNO As String, ByVal lng病人ID As Long, _
    Optional ByVal blnFirstNO As Boolean) As Boolean
    '功能：补打挂号票据，病人收费管理中的按病人补打票据调用，多张单据时会分别调用
    '入参：
    '   blnFirstNO:是否第一张单据
    '返回：打印成功返回True，打印失败返回False

    Dim blnStartFactUseType  As Boolean, strUseType As String
    Dim intInsure As Integer, blnVirtualPrint As Boolean
    
    If blnFirstNO Then
        Set gfrmMain = frmMain
        If Not CheckValid Then Exit Function
        
        Set gcnOracle = cnMain
        glngSys = lngSys
        gstrDBUser = strDbUser
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrUnitName = GetUnitName
        If GetUserInfo() = False Then Exit Function
        
        gstrPrivs = ";" & GetPrivFunc(lngSys, 1111) & ";"
        If zlStr.IsHavePrivs(gstrPrivs, "补打票据") = False Then
            MsgBox "你没有被授予使用挂号的补打票据功能的权限,如需使用,请与系统管理员联系!", vbInformation, gstrSysName
            Exit Function
        End If
        
        Call InitLocPar(1111)
        Call InitSysPar
        Call gclsInsure.InitOracle(gcnOracle)
    End If
    
    intInsure = ExistInsure(strNO)
    If intInsure <> 0 Then
        blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure)
    End If
    
    If gblnSharedInvoice Then
        '挂号用门诊票据:42703
        blnStartFactUseType = zlStartFactUseType("1")
        If blnStartFactUseType Then
            strUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
        End If
    End If
    '第一张单据时需要确认发票号，其它不用确认发票号
    If Not RePrintBill(frmMain, 4, strNO, 0, 0, blnVirtualPrint, strUseType, True, blnFirstNO) Then Exit Function
    PrintRegistBill = True
End Function

Public Function Register(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, _
    ByVal strDbUser As String, ByVal blnSharedInvoice As Boolean, ByVal lngInvoiceGroupID As Long) As Boolean
'功能：收费直接调用挂号功能
'返回：是否挂过号
    Static blnNoFirst As Boolean  '是否第一次调用,初始化全局变量
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, 1111) & ";"
    If Not blnNoFirst Then
        blnNoFirst = True
        Set gfrmMain = frmMain
        If Not CheckValid Then blnNoFirst = False: Exit Function
        
        Set gcnOracle = cnMain
        glngSys = lngSys
        gstrDBUser = strDbUser
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrUnitName = GetUnitName
        
        If InStr(";" & gstrPrivs & ";", ";挂收费号;") = 0 And InStr(";" & gstrPrivs & ";", ";挂免费号;") = 0 Then
            MsgBox "你没有被授予使用挂号功能的权限,如需使用,请与系统管理员联系!", vbInformation, gstrSysName
            blnNoFirst = False
            Exit Function
        End If
        
        '医保接口
        Call gclsInsure.InitOracle(gcnOracle)
        '用户信息
        Call GetUserInfo
        '初始化参数
        Call InitLocPar(1111)
        Call InitSysPar
        Call CreateSquareCardObject(gfrmMain, 1111)
        Call CreatePublicExpenseObject(1111)
    End If
    gblnSharedInvoice = blnSharedInvoice
    If gblnSharedInvoice Then glng挂号ID = lngInvoiceGroupID
    
    '创建第三方票据打印部件
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, 1111, UserInfo.编号, UserInfo.姓名)
    End If
    
    '调用挂号
    Err.Clear
    On Error Resume Next
    If gbytRegistMode = 0 Then
        If gbln精简界面 Then
            frmRegistEditSimple.mblnCharge = True
            frmRegistEditSimple.mlngModul = 1111
            frmRegistEditSimple.mstrPrivs = gstrPrivs
            frmRegistEditSimple.mbytMode = 0
            frmRegistEditSimple.mbytInState = 0
            If gfrmMain Is Nothing Then
                frmRegistEditSimple.Show
            Else
                frmRegistEditSimple.Show 1, gfrmMain
            End If
        Else
            frmRegistEdit.mblnCharge = True
            frmRegistEdit.mlngModul = 1111
            frmRegistEdit.mstrPrivs = gstrPrivs
            frmRegistEdit.mbytMode = 0
            frmRegistEdit.mbytInState = 0
            If gfrmMain Is Nothing Then
                frmRegistEdit.Show
            Else
                frmRegistEdit.Show 1, gfrmMain
            End If
        End If
    Else
        If Sys.Currentdate < gdatRegistTime Then
            If gbln精简界面 Then
                frmRegistEditSimple.mblnCharge = True
                frmRegistEditSimple.mlngModul = 1111
                frmRegistEditSimple.mstrPrivs = gstrPrivs
                frmRegistEditSimple.mbytMode = 0
                frmRegistEditSimple.mbytInState = 0
                If gfrmMain Is Nothing Then
                    frmRegistEditSimple.Show
                Else
                    frmRegistEditSimple.Show 1, gfrmMain
                End If
            Else
                frmRegistEdit.mblnCharge = True
                frmRegistEdit.mlngModul = 1111
                frmRegistEdit.mstrPrivs = gstrPrivs
                frmRegistEdit.mbytMode = 0
                frmRegistEdit.mbytInState = 0
                If gfrmMain Is Nothing Then
                    frmRegistEdit.Show
                Else
                    frmRegistEdit.Show 1, gfrmMain
                End If
            End If
        Else
            frmRegistEditNew.mblnCharge = True
            frmRegistEditNew.mlngModul = 1111
            frmRegistEditNew.mstrPrivs = gstrPrivs
            frmRegistEditNew.mbytMode = 0
            frmRegistEditNew.mbytInState = 0
            If gfrmMain Is Nothing Then
                frmRegistEditNew.Show
            Else
                frmRegistEditNew.Show 1, gfrmMain
            End If
        End If
    End If
    Register = gblnOk
    
    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
End Function

Public Function StationRegister(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, ByVal strRoom As String, _
    Optional ByVal blnNoneCut As Boolean, Optional ByVal blnPrice As Boolean = True, _
    Optional lng病人ID As Long = 0, Optional strDbUser As String) As String
'功能：门诊医生站调用挂号,不处理授权
'参数：strRoom=接诊诊室
'      blnNoneCut=是否不允许使用打折费别
'      blnPrice=是否允许挂号费用生成划价单收费
'返回：如果挂号成功，则返回挂号单号，否则返回空
    
    Static blnNFStationRegister As Boolean  '是否第一次调用,初始化全局变量
   
   gstrPrivsStation = ";" & GetPrivFunc(lngSys, 1260) & ";" '56274
    If Not blnNFStationRegister Then
        blnNFStationRegister = True
    
        If Not CheckValid Then Exit Function
    
        glngSys = lngSys
        Set gcnOracle = cnMain
        Set gfrmMain = frmMain
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrUnitName = GetUnitName
        gstrDBUser = strDbUser  '52689
        
        '用户信息
        If Not GetUserInfo Then
            MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
        '医保接口
        Call gclsInsure.InitOracle(gcnOracle)
        '初始化参数
        Call InitLocPar(1111)
        Call InitSysPar
        Call CreateSquareCardObject(gfrmMain, 1111)
        Call CreatePublicExpenseObject(1111)
    End If
    '需要每次赋值,避免因进入其它模块而改变
    gstrPrivs = ";挂号;挂免费号;建立病案;退号;" '固有权限,调用程序脚本已授权
    
    '调用挂号
    Err.Clear
    On Error Resume Next
    frmRegistEdit.mblnStation = True
    frmRegistEdit.mstrRoom = strRoom
    frmRegistEdit.mstrRegNo = ""
    frmRegistEdit.mblnNoneCut = blnNoneCut
    frmRegistEdit.mblnStationPrice = blnPrice
    frmRegistEdit.mlngModul = 1111
    frmRegistEdit.mstrPrivs = gstrPrivs
    frmRegistEdit.mbytMode = 0
    frmRegistEdit.mbytInState = 0
    If lng病人ID <> 0 Then
        frmRegistEdit.zl_StationInitPatient (lng病人ID)
    End If
    If gfrmMain Is Nothing Then
        frmRegistEdit.Show
    Else
        frmRegistEdit.Show 1, frmMain
    End If
    If gblnOk Then
        StationRegister = frmRegistEdit.mstrRegNo
    End If
End Function

Public Function StationBespeak(frmMain As Object, cnMain As ADODB.Connection, _
    lngSys As Long, ByVal strRoom As String, Optional ByVal blnNoneCut As Boolean, _
    Optional ByVal lng病人ID As Long = 0, Optional strDbUser As String) As String
    '功能：门诊医生站调用预约挂号,不处理授权
    '参数：strRoom=接诊诊室
    '      blnNoneCut=是否不允许使用打折费别
    '返回：如果挂号成功，则返回挂号单号，否则返回空
    Static blnNFStationBespeak As Boolean  '是否第一次调用,初始化全局变量
   gstrPrivsStation = ";" & GetPrivFunc(lngSys, 1260) & ";" '56274
    If Not blnNFStationBespeak Then
        blnNFStationBespeak = True
        If Not CheckValid Then Exit Function
        
        glngSys = lngSys
        Set gcnOracle = cnMain
        Set gfrmMain = frmMain
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrUnitName = GetUnitName
        
        gstrDBUser = strDbUser  '52689
        '用户信息
        If Not GetUserInfo Then
            MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '医保接口
        Call gclsInsure.InitOracle(gcnOracle)
            
        '初始化参数
        Call InitLocPar(1111)
        Call InitSysPar
        
        Call CreateSquareCardObject(gfrmMain, 1111)
        Call CreatePublicExpenseObject(1111)
    End If
    '需要每次赋值,避免因进入其它模块而改变
    gstrPrivs = ";预约挂号;挂免费号;建立病案;" '固有权限,调用程序脚本已授权
    
    '调用挂号
    Err.Clear
    On Error Resume Next
    frmRegistEdit.mblnStation = True
    frmRegistEdit.mstrRoom = strRoom
    frmRegistEdit.mstrRegNo = ""
    frmRegistEdit.mblnNoneCut = blnNoneCut
    frmRegistEdit.mlngModul = 1111
    frmRegistEdit.mstrPrivs = gstrPrivs
    frmRegistEdit.mbytMode = 1
    frmRegistEdit.mbytInState = 0
    If lng病人ID <> 0 Then
        frmRegistEdit.zl_StationInitPatient (lng病人ID)
    End If
    If gfrmMain Is Nothing Then
        frmRegistEdit.Show
    Else
        frmRegistEdit.Show 1, frmMain
    End If
    If gblnOk Then
        StationBespeak = frmRegistEdit.mstrRegNo
    End If
End Function
Public Sub zlPrintBespeak(ByVal frmMain As Object, ByVal cnMain As ADODB.Connection, _
    ByVal lngSys As Long, ByVal strDbUser As String, ByVal strPrivs As String, ByVal strNO As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印预约挂号单
    '入参:strNO=预约挂号单号
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-12-24 10:34:14
    '说明:
    '问题:56274
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If InStr(1, strPrivs, ";预约挂号单;") = 0 Then Exit Sub
    Call ReportOpen(cnMain, lngSys, "ZL" & lngSys \ 100 & "_BILL_1111_1", frmMain, "NO=" & strNO, 2)
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

Private Sub Class_Terminate()
    CloseSquareCardObject
End Sub
