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

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名

    '------------------------------------------------
On Error GoTo errHandle
    
    Set gcnOracle = cnMain
    Set gfrmMain = frmMain
    gstrDBUser = strDBUser
    glngModul = lngModul
    glngSys = lngSys

    '读取注册表，判断是否启用调试日志
    gblnUseDebugLog = IIf(Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPacsMain", "UseDebugLog", 0)) = 0, False, True)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPacsMain", "UseDebugLog", IIf(gblnUseDebugLog, 1, 0))
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub
    
    
    If Not InitSysPar Then Exit Sub '初始化系统参数
    
    Select Case lngModul
        Case 1056 '影像检查项目
            frmRadLists.Show , gfrmMain
        Case 1057 '影像设备目录
            frmPACSDevice.Show , gfrmMain
        Case 1079
            frmPatholSpecimenCfg.Show , gfrmMain
        Case 1290 '影像技师工作站
            Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '初始化医嘱，病历核心部件
            Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)

            If mfrmPacsMain Is Nothing Then
                Set mfrmPacsMain = New frmPacsMain
            End If
            
            mfrmPacsMain.Caption = "影像医技工作站"
            mfrmPacsMain.ShowStation lngModul, gfrmMain
        
        Case 1291   '影像采集工作站
            
            Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '初始化医嘱，病历核心部件
            Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
            
            If mfrmVideoMain Is Nothing Then
                Set mfrmVideoMain = New frmPacsMain
            End If
            
            mfrmVideoMain.Caption = "影像采集工作站"
            mfrmVideoMain.ShowStation lngModul, gfrmMain
            
        Case 1292   '影像流程控制
            frmWorkFlow.Show , gfrmMain
        Case 1294   '影像病理工作站
            Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '初始化医嘱，病历核心部件
            Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
            
            If mfrmPatholMain Is Nothing Then
                Set mfrmPatholMain = New frmPacsMain
            End If
            
            mfrmPatholMain.Caption = "病理检查工作站"
            mfrmPatholMain.ShowStation lngModul, gfrmMain
            
        Case 1295   '病理归档工作站
            frmPatholArchivesStation.Show , gfrmMain
        Case 1296   '病理借还工作站
            frmPatholReborrowStation.Show , gfrmMain
        Case 1297   '病理材料遗失处理工作站
            frmPatholLoseStation.Show , gfrmMain
    End Select

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名

    '------------------------------------------------
On Error GoTo errHandle
    Set gcnOracle = cnMain
    Set gfrmMain = Nothing
    gstrDBUser = strDBUser
    glngModul = lngModul
    glngSys = lngSys

    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub

    If Not InitSysPar Then Exit Sub '初始化系统参数

    Select Case lngModul
        Case 1056 '影像检查项目
            zlCommFun.ShowChildWindow frmRadLists.hWnd, lngMain
        Case 1057 '影像设备目录
            zlCommFun.ShowChildWindow frmPACSDevice.hWnd, lngMain
        Case 1057 '影像设备目录
            zlCommFun.ShowChildWindow frmPACSDevice.hWnd, lngMain
        Case 1290 '影像技师工作站
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs) '初始化医嘱，病历核心部件
            Call gobjRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)

            If mfrmPacsMain Is Nothing Then
                Set mfrmPacsMain = New frmPacsMain
                mfrmPacsMain.Caption = "影像医技工作站"
            End If
            
            zlCommFun.ShowChildWindow mfrmPacsMain.hWnd, lngMain, True
            mfrmPacsMain.ShowStation lngModul, gfrmMain     '此时gfrmMain为nothing
            
        Case 1291   '影像采集工作站
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs) '初始化医嘱，病历核心部件
            Call gobjRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)

            If mfrmVideoMain Is Nothing Then
                Set mfrmVideoMain = New frmPacsMain
                mfrmVideoMain.Caption = "影像采集工作站"
            End If
            
            zlCommFun.ShowChildWindow mfrmVideoMain.hWnd, lngMain, True
            mfrmVideoMain.ShowStation lngModul, gfrmMain    '此时gfrmMain为nothing
            
        Case 1292   '影像流程控制
            zlCommFun.ShowChildWindow frmWorkFlow.hWnd, lngMain
        Case 1294   '影像病理工作站
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs) '初始化医嘱，病历核心部件
            Call gobjRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)

            If mfrmPatholMain Is Nothing Then
                Set mfrmPatholMain = New frmPacsMain
                mfrmPatholMain.Caption = "病理检查工作站"
            End If
            
            zlCommFun.ShowChildWindow mfrmPatholMain.hWnd, lngMain, True
            mfrmPatholMain.ShowStation lngModul, gfrmMain   '此时gfrmMain为nothing
            
        Case 1295   '病理归档工作站
            zlCommFun.ShowChildWindow frmPatholArchivesStation.hWnd, lngMain, True
        Case 1296   '病理借还工作站
            zlCommFun.ShowChildWindow frmPatholReborrowStation.hWnd, lngMain, True
        Case 1297   '病理材料遗失处理工作站
            zlCommFun.ShowChildWindow frmPatholLoseStation.hWnd, lngMain, True
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
On Error GoTo errHandle
    
    Dim frmThis As Form
    
    For Each frmThis In Forms
        UnLoadForm frmThis
    Next
    
    CloseWindows = (Forms.Count = 0)

errHandle:
End Function

Private Sub UnLoadForm(objForm As Object)
On Error GoTo errHandle
    Unload objForm
Exit Sub
errHandle:
    Debug.Print err.Description
End Sub

Private Sub Class_Initialize()
    VERSION = App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
On Error GoTo errHandle
    Unload frmPacsMain
    Set frmPacsMain = Nothing
    
errHandle:
End Sub
