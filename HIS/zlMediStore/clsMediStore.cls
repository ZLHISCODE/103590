VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMediStore"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit
  Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, FrmMain As Object, strDbUser As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngSys : 系统编号
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '返回：
    '------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strDept As String
    Dim strPrivs As String
    
    Set gcnOracle = cnMain
    
    glngSys = lngSys
    
'    Set frmWin = FrmMain
    gstrAviPath = GetSetting("ZLSOFT", "注册信息", "gstrAviPath", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrVersion = GetSetting("ZLSOFT", "注册信息", "gstrVersion", "")
    
    gstrMatchMethod = GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0)
    gbytSimpleCodeTrans = Val(zldatabase.GetPara("简码匹配方式切换", , , 1))
    
    glngModul = lngModul
    gstrprivs = GetPrivFunc(lngSys, lngModul)
'    zlbillprint.BillInit gcnOracle
    
     '获取登录用户信息
    gstrDbUser = strDbUser   '  GetSetting("ZLHIS", "登录信息", frmMain.hwnd, GetSetting("ZLHIS", "登录信息", "USER", ""))
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not CheckValid Then Exit Sub
    
    '读取药品流通部分要使用的系统参数值
    Call GetSysParms
    
    '取药品价格等级
    Call GetPriceClass
    
    '-------------------------------------------------
    '1300  药品外购入库管理
    '1301  药品自制入库管理
    '1302  药品其他入库管理
    '1303  药品库存差价调整管理
    '1304  药品移库管理
    '1305  药品领用管理
    '1306  药品其他出库管理
    '1307  药品盘点管理
    '1308  药品差价重算
    '1309  药品库存查询
    '1320  药品付款管理
    
    '1330  药品计划管理
    '1331  药品质量管理
    '1332  药品结存管理
    '1333  药品调价管理
    '2287  药品入库验收
    
    Dim frmPurchaseList As Form
    Dim frmTemp As Form
    Dim strTitle As String
    
    
    On Error Resume Next
    
    Call setNOtExcetePrice '先执行到时间还没有执行的价格
    Select Case lngModul
        Case 1300
            strTitle = "药品外购入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
            
             '.Show , frmMain            '药品外购入库管理
        Case 1301
            strTitle = "药品自制入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        Case 1302                        '药品其他入库管理
            strTitle = "药品其他入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
'        Case 1303
'            strTitle = "库存差价调整管理"
'            For Each frmTemp In Forms
'                If frmTemp.Caption = strTitle Then
'                    frmTemp.Show
'                    frmTemp.ZOrder 0
'                    Exit Sub
'                End If
'            Next
'            Set frmPurchaseList = New frmMainList
'            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        Case 1304
            strTitle = "药品移库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        
        Case 1305
            strTitle = "药品领用管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        Case 1306
            strTitle = "药品其他出库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        Case 1307
            strTitle = "药品盘点管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmNewCheckMain
            frmPurchaseList.ShowList lngModul, strTitle, FrmMain
        Case 1320
            'frmDrugPaymentList.Show , FrmMain
        Case 1321
            'frm应付款查询.Show , FrmMain
        Case 1330
            frmDrugPlanList.Show , FrmMain
        Case 1331
            frmDrugQualityList.Show , FrmMain
'        Case 1308
'            frmDiffPriceRecalCard.Show , FrmMain
        Case 1309
            frmDrugQuery.Show , FrmMain
        Case 1332
            frmClosingAccount.Show , FrmMain
        Case 1333
            frmMediPrice.Show , FrmMain
        Case 1348 '药品验收入库
            frmMediCheckList.Show , FrmMain
    End Select
End Sub

Private Sub setNOtExcetePrice()
    '将到时间还未执行调价药品执行调价
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct i.Id As 药品id " & _
               " From 收费项目目录 I, 收费价目 N, 药品规格 P" & _
               " Where i.Id = n.收费细目id And i.Id = p.药品id And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & _
                   " n.变动原因 = 0 And Sysdate>n.执行日期" & GetPriceClassString("N") & _
               " Union " & _
               " Select Distinct a.药品id From 药品价格记录 A Where a.记录状态 = 0 And a.执行日期 <= Sysdate " & _
               " Order By 药品id "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "执行调价")
    
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    For i = 0 To rsTemp.RecordCount - 1
        gstrSQL = "Zl_药品收发记录_Adjust(" & rsTemp!药品ID & ")"
        zldatabase.ExecuteProcedure gstrSQL, "执行调价"
        rsTemp.MoveNext
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, strDbUser As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngSys : 系统编号
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   lngMain:主窗体句柄
    '返回：
    '------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strDept As String
    Dim strPrivs As String
    
    Set gcnOracle = cnMain
    
    glngSys = lngSys
    
    gstrAviPath = GetSetting("ZLSOFT", "注册信息", "gstrAviPath", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrVersion = GetSetting("ZLSOFT", "注册信息", "gstrVersion", "")
    
    gstrMatchMethod = GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0)
    gbytSimpleCodeTrans = Val(zldatabase.GetPara("简码匹配方式切换", , , 1))
    
    glngModul = lngModul
    gstrprivs = GetPrivFunc(lngSys, lngModul)
    
    '获取登录用户信息
    gstrDbUser = strDbUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not CheckValid Then Exit Sub
    
    '读取药品流通部分要使用的系统参数值
    Call GetSysParms
   
    '-------------------------------------------------
    '1300  药品外购入库管理
    '1301  药品自制入库管理
    '1302  药品其他入库管理
    '1303  药品库存差价调整管理
    '1304  药品移库管理
    '1305  药品领用管理
    '1306  药品其他出库管理
    '1307  药品盘点管理
    '1308  药品差价重算
    '1309  药品库存查询
    '1320  药品付款管理
    
    '1330  药品计划管理
    '1331  药品质量管理
    '1332  药品结存管理
    '1333  药品调价管理
    
    Dim frmPurchaseList As Form
    Dim frmTemp As Form
    Dim strTitle As String
    
    On Error Resume Next
    
    Call setNOtExcetePrice '先执行到时间还没有执行的价格
    Select Case lngModul
        Case 1300 '药品外购入库管理
            strTitle = "药品外购入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1301
            strTitle = "药品自制入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1302                        '药品其他入库管理
            strTitle = "药品其他入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
'        Case 1303
'            strTitle = "库存差价调整管理"
'            For Each frmTemp In Forms
'                If frmTemp.Caption = strTitle Then
'                    frmTemp.Show
'                    frmTemp.ZOrder 0
'                    Exit Sub
'                End If
'            Next
'            Set frmPurchaseList = New frmMainList
'            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1304
            strTitle = "药品移库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        
        Case 1305
            strTitle = "药品领用管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1306
            strTitle = "药品其他出库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmMainList
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1307
            strTitle = "药品盘点管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmPurchaseList = New frmNewCheckMain
            frmPurchaseList.ShowList lngModul, strTitle, lngMain
        Case 1320
            'OS.ShowChildWindow frmDrugPaymentList.hWnd, lngMain
        Case 1321
            'OS.ShowChildWindow frm应付款查询.hWnd, lngMain
        Case 1330
            OS.ShowChildWindow frmDrugPlanList.hWnd, lngMain
            
        Case 1331
            OS.ShowChildWindow frmDrugQualityList.hWnd, lngMain
'        Case 1308
'            OS.ShowChildWindow frmDiffPriceRecalCard.hWnd, lngMain
        Case 1309
            OS.ShowChildWindow frmDrugQuery.hWnd, lngMain
        Case 1332
            OS.ShowChildWindow frmClosingAccount.hWnd, lngMain
        Case 1333
            OS.ShowChildWindow frmMediPrice.hWnd, lngMain
        Case 1348 '药品验收入库
            OS.ShowChildWindow frmMediCheckList.hWnd, lngMain
    End Select
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.count = 0)
End Function

