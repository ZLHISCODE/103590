VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Version As String

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    Call RegBillFile
End Sub

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String, Optional varParameter As Variant)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '返回：
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting("ZLSOFT", "注册信息", UCase("gstrAviPath"), "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", UCase("gstrSysName"), "")
    gstrVersion = GetSetting("ZLSOFT", "注册信息", UCase("gstrVersion"), "")
    gstr单位名称 = GetUnitName()
    gstrDbUser = strDbUser
    glngSys = lngSys
    glngModul = lngModul
    
    Set gfrmMain = frmMain
    glngMain = 0
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Select Case lngModul
        Case 1031 '记帐单设计
            gstrPrivs = GetPrivFunc(lngSys, lngModul)
            
            Set frmDesign = Nothing
            frmDesign.Show , frmMain
        Case -1
            gblnOK = False
            
            '参数含义依次是：
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、gstrPrivs、blnViewCancel
            gstrPrivs = varParameter(7)
            
            '医保接口(住院记帐)
            If varParameter(1) <> 3 Then
                Call gclsInsure.InitOracle(gcnOracle)
            End If
            
            If InStr(1, GetInsidePrivs(Enum_Inside_Program.p记帐操作), "专项记帐") = 0 Then
                MsgBox "你没有操作专项记帐单的权限!", vbInformation, gstrSysName
                varParameter = gblnOK
                Exit Sub
            End If
            
            With frmCharge
                .mlng记帐ID = Val(varParameter(0))
                .mbytUseType = varParameter(1)
                .mbytInState = varParameter(2)
                .mstrInNO = varParameter(3)
                .mlngUnitID = Val(varParameter(4))
                .mlngDeptID = Val(varParameter(5))
                .mlng病人ID = Val(varParameter(6))
                .mstrPrivs = gstrPrivs
                
                If .mbytInState = sta查阅 Then
                    .mblnViewCancel = varParameter(8)
                Else
                    .mblnViewCancel = False
                End If
            End With
            Call InitLocPar(CByte(varParameter(1)))
            Call InitSysPar
            frmCharge.MainProc
            varParameter = gblnOK
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String, Optional varParameter As Variant)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   lngMain:主窗体句柄
    '返回：
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting("ZLSOFT", "注册信息", UCase("gstrAviPath"), "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", UCase("gstrSysName"), "")
    gstrVersion = GetSetting("ZLSOFT", "注册信息", UCase("gstrVersion"), "")
    gstr单位名称 = GetUnitName()
    gstrDbUser = strDbUser
    glngSys = lngSys
    glngModul = lngModul
    
    Set gfrmMain = Nothing
    glngMain = lngMain
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Select Case lngModul
        Case 1031 '记帐单设计
            gstrPrivs = GetPrivFunc(lngSys, lngModul)
            
            Set frmDesign = Nothing
            zlCommFun.ShowChildWindow frmDesign.hwnd, lngMain
        Case -1
            gblnOK = False
            
            '参数含义依次是：
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、gstrPrivs、blnViewCancel
            gstrPrivs = varParameter(7)
            
            '医保接口(住院记帐)
            If varParameter(1) <> 3 Then
                Call gclsInsure.InitOracle(gcnOracle)
            End If
            
            If InStr(1, GetInsidePrivs(Enum_Inside_Program.p记帐操作), "专项记帐") = 0 Then
                MsgBox "你没有操作专项记帐单的权限!", vbInformation, gstrSysName
                varParameter = gblnOK
                Exit Sub
            End If
            
            With frmCharge
                .mlng记帐ID = Val(varParameter(0))
                .mbytUseType = varParameter(1)
                .mbytInState = varParameter(2)
                .mstrInNO = varParameter(3)
                .mlngUnitID = Val(varParameter(4))
                .mlngDeptID = Val(varParameter(5))
                .mlng病人ID = Val(varParameter(6))
                .mstrPrivs = gstrPrivs
                
                If .mbytInState = sta查阅 Then
                    .mblnViewCancel = varParameter(8)
                Else
                    .mblnViewCancel = False
                End If
            End With
            Call InitLocPar(CByte(varParameter(1)))
            Call InitSysPar
            frmCharge.MainProc
            varParameter = gblnOK
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
