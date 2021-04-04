VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPACSWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public VERSION As String

'中药配方界面所需要的病人信息
Public Type TYPE_PatiInfoEx
        lng病人ID As Long
        lng主页ID As Long '主页ID
        str挂号单 As String '挂号单号
        str性别 As String  '病人性别
        int婴儿 As Integer
        int险类 As Integer '医保病人的险类
        bln医保 As Boolean '是否医保病人
        lng病人科室ID As Long '用于确定中药配方的缺省药房
End Type

Private mfrmPacsMain As frmPacsMain
Private mfrmVideoMain As frmPacsMain
Private mfrmPatholMain As frmPacsMain
Private mobjPacsInterface As Object


Private Sub OpenNewRichReportModul(ByVal lngSys As Long, ByVal lngModul As Long, lngMain As Long)
'打开新版报告文档编辑器各管理模块
    Dim objRegister As Object
    
    If mobjPacsInterface Is Nothing Then
        Set mobjPacsInterface = DynamicCreate("ZLSoft.ZLPacs.Interface.PacsInterface", "PACS智能报告编辑器接口部件zlPacsInterfaceC")
    End If
    
    If Not mobjPacsInterface Is Nothing Then
        If InitRegister = False Then Exit Sub
        
        BUGEX "zlInit InitRegister " & gstrUserPswd
        
        Call mobjPacsInterface.CodeMain(glngSys, lngModul, "", lngMain, gstrUserName, gstrUserPswd, gstrServerName)
    End If
End Sub


Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名

    '------------------------------------------------
On Error GoTo ErrHandle
    
    Set gcnOracle = cnMain
    Set gfrmMain = frmMain
    
    glngMainHwnd = gfrmMain.hWnd
    gstrDBUser = strDBUser
    gstrInputPwd = frmMain.objLogin.InputPwd
    glngModul = lngModul
    glngSys = lngSys

    '读取注册表，判断是否启用调试日志
    gblnUseDebugLog = IIf(Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPacsMain", "UseDebugLog", 0)) = 0, False, True)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPacsMain", "UseDebugLog", IIf(gblnUseDebugLog, 1, 0))
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub
    
    Call WriteLog("CodeMan -> Step 1")
    
    If Not InitSysPar Then Exit Sub '初始化系统参数
    
    Select Case lngModul
        Case 1049, 2280, 2281, 2282, 2283, 2284, 2286
            Call OpenNewRichReportModul(lngSys, lngModul, frmMain.hWnd)
        
        Case 1056 '影像检查项目
            frmRadLists.Show , gfrmMain
        Case 1057 '影像设备目录
            frmPACSDevice.Show , gfrmMain
        Case 1079
            frmPatholSpecimenCfg.Show , gfrmMain
        Case 1290 '影像技师工作站
                        
            gstrOracleOwner = zlDatabase.GetPara("XWOracle拥有者", glngSys, G_LNG_XWPACSVIEW_MODULE, "zlhis")
            
            If mfrmPacsMain Is Nothing Then
                Set mfrmPacsMain = CreateMainModule(gfrmMain, "影像医技工作站", False)
            Else
                If mfrmPacsMain.IsValid = False Then
                    Set mfrmPacsMain = CreateMainModule(gfrmMain, "影像医技工作站", False)
                Else
                    mfrmPacsMain.WindowState = 0
                    mfrmPacsMain.Show , gfrmMain
                End If
            End If
        
            Set gfrmPacsMain = mfrmPacsMain
        Case 1291   '影像采集工作站

            If mfrmVideoMain Is Nothing Then
                Set mfrmVideoMain = CreateMainModule(gfrmMain, "影像采集工作站", False)
            Else
                If mfrmVideoMain.IsValid = False Then
                    Set mfrmVideoMain = CreateMainModule(gfrmMain, "影像采集工作站", False)
                Else
                    mfrmVideoMain.WindowState = 0
                    mfrmVideoMain.Show , gfrmMain
                End If
            End If
            
'        Case 1292   '影像流程控制
'            frmWorkFlow.Show , gfrmMain
        Case 1294   '影像病理工作站
            If mfrmPatholMain Is Nothing Then
                Set mfrmPatholMain = CreateMainModule(gfrmMain, "病理检查工作站", False)
            Else
                If mfrmPatholMain.IsValid = False Then
                    Set mfrmPatholMain = CreateMainModule(gfrmMain, "病理检查工作站", False)
                Else
                    mfrmPatholMain.WindowState = 0
                    mfrmPatholMain.Show , gfrmMain
                End If
            End If
            
        Case 1295   '病理归档工作站
            frmPatholArchivesStation.Show , gfrmMain
        Case 1296   '病理借还工作站
            frmPatholReborrowStation.Show , gfrmMain
        Case 1297   '病理材料遗失处理工作站
            frmPatholLoseStation.Show , gfrmMain
    End Select
    
    Call WriteLog("CodeMan -> Step 2")
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function CreateMainModule(ByVal objParent As Object, ByVal strCaption As String, ByVal blnIsBHMain As Boolean) As frmPacsMain
'创建主模块
    Call WriteLog("CreateMainModule -> Step 1： 初始化基础部件.")
    Call InitBaseComponent(objParent, blnIsBHMain)

    Call WriteLog("CreateMainModule -> Step 2： 开始创建显示工作模块 [" + strCaption + "].")
    Set CreateMainModule = New frmPacsMain
    
    If blnIsBHMain = True Then
        OS.ShowChildWindow CreateMainModule.hWnd, glngMainHwnd, True
        CreateMainModule.ShowStation glngModul, Nothing
    Else
        CreateMainModule.ShowStation glngModul, objParent
    End If

    CreateMainModule.Caption = strCaption
        
    Call WriteLog("CreateMainModule -> Step 3： 工作模块 [" + strCaption + "] 创建完成.")
End Function

Private Sub InitBaseComponent(ByVal objParent As Object, ByVal blnIsBHMain As Boolean)
'初始化基础部件
'blnIsBHMain=false时objParent为frmMain对象
'blnIsBHMain=true时objParent为objLogin对象

    Set gobjEmr = objParent.mobjEMR
    
    If blnIsBHMain = False Then
        Call gobjKernel.InitCISKernel(gcnOracle, objParent, glngSys, gstrPrivs) '初始化医嘱，病历核心部件
        Call gobjRichEPR.InitRichEPR(gcnOracle, objParent, glngSys, False)

        Call InitPacsControlWithInPro(gcnOracle, glngMainHwnd, glngSys, glngModul, objParent, gblnUseDebugLog)
    Else
        Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs) '初始化医嘱，病历核心部件
        Call gobjRichEPR.InitRichEPR(gcnOracle, objParent, glngSys, False)
        
        Call InitPacsControlWithInPro(gcnOracle, glngMainHwnd, glngSys, glngModul, Nothing, gblnUseDebugLog)
    End If
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal objLogin As Object)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名

    '------------------------------------------------
On Error GoTo ErrHandle
    Set gcnOracle = cnMain
    Set gfrmMain = Nothing
    
    glngMainHwnd = lngMain
    gstrDBUser = strDBUser
    
    If Not objLogin Is Nothing Then
        gstrInputPwd = objLogin.InputPwd
    End If
    
    glngModul = lngModul
    glngSys = lngSys

    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub

    If Not InitSysPar Then Exit Sub '初始化系统参数
    
    '读取注册表，判断是否启用调试日志
    gblnUseDebugLog = IIf(Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPacsMain", "UseDebugLog", 0)) = 0, False, True)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPacsMain", "UseDebugLog", IIf(gblnUseDebugLog, 1, 0))
        

    Select Case lngModul
        Case 1049, 2280, 2281, 2282, 2283, 2284, 2286
            Call OpenNewRichReportModul(lngSys, lngModul, lngMain)
        
        Case 1056 '影像检查项目
            OS.ShowChildWindow frmRadLists.hWnd, lngMain
        Case 1057 '影像设备目录
            OS.ShowChildWindow frmPACSDevice.hWnd, lngMain
        Case 1079
            OS.ShowChildWindow frmPatholSpecimenCfg.hWnd, lngMain
        Case 1290 '影像技师工作站
            
            gstrOracleOwner = zlDatabase.GetPara("XWOracle拥有者", glngSys, G_LNG_XWPACSVIEW_MODULE, "zlhis")
                        
            If mfrmPacsMain Is Nothing Then
                Set mfrmPacsMain = CreateMainModule(objLogin, "影像医技工作站", True)
            Else
                If mfrmPacsMain.IsValid = False Then
                    Set mfrmPacsMain = CreateMainModule(objLogin, "影像医技工作站", True)
                Else
                    mfrmPacsMain.WindowState = 0
                    mfrmPacsMain.Show , Nothing
                End If
            End If
            
            Set gfrmPacsMain = mfrmPacsMain
        Case 1291   '影像采集工作站
            
            If mfrmVideoMain Is Nothing Then
                Set mfrmVideoMain = CreateMainModule(objLogin, "影像采集工作站", True)
            Else
                If mfrmVideoMain.IsValid = False Then
                    Set mfrmVideoMain = CreateMainModule(objLogin, "影像采集工作站", True)
                Else
                    mfrmVideoMain.WindowState = 0
                    mfrmVideoMain.Show , Nothing
                End If
            End If
            
            
'        Case 1292   '影像流程控制
'            zlCommFun.ShowChildWindow frmWorkFlow.hWnd, lngMain
        Case 1294   '影像病理工作站
            If mfrmPatholMain Is Nothing Then
                Set mfrmPatholMain = CreateMainModule(objLogin, "病理检查工作站", True)
            Else
                If mfrmPatholMain.IsValid = False Then
                    Set mfrmPatholMain = CreateMainModule(objLogin, "病理检查工作站", True)
                Else
                    mfrmPatholMain.WindowState = 0
                    mfrmPatholMain.Show , Nothing
                End If
            End If
            
        Case 1295   '病理归档工作站
            OS.ShowChildWindow frmPatholArchivesStation.hWnd, lngMain, True
        Case 1296   '病理借还工作站
            OS.ShowChildWindow frmPatholReborrowStation.hWnd, lngMain, True
        Case 1297   '病理材料遗失处理工作站
            OS.ShowChildWindow frmPatholLoseStation.hWnd, lngMain, True
    End Select
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
On Error GoTo ErrHandle
    
    Dim frmThis As Form
    
    For Each frmThis In Forms
        UnLoadForm frmThis
    Next
    
    Call gobjEmr.CloseForms
    CloseWindows = (Forms.Count = 0)

ErrHandle:
End Function

Private Sub UnLoadForm(objForm As Object)
On Error GoTo ErrHandle
    Unload objForm
Exit Sub
ErrHandle:
    Debug.Print err.Description
End Sub

Private Sub Class_Initialize()
    VERSION = App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrHandle
    Unload frmPacsMain
    Set frmPacsMain = Nothing
    Set mobjPacsInterface = Nothing
ErrHandle:
End Sub
