VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChargeBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmChargeBillTotal As frmChargeBillTotal
Public Enum TotalType
    EM_收费员轧帐 = 1
    EM_小组收款 = 2
    EM_小组轧帐 = 3
    EM_财务收款 = 4
End Enum

Private Sub initChargeBillTotalWin()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收款及票据汇总
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-03 18:22:50
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmChargeBillTotal Is Nothing Then
        Set mfrmChargeBillTotal = New frmChargeBillTotal
        Load mfrmChargeBillTotal
    End If
End Sub

Public Function InitOracle(cnMain As ADODB.Connection, ByVal strDbUser As String, lngSys As Long) As Boolean
    '-------------------------------------------------------------------------------------------------
    '功能:初始化连接
    '入参: strDBUser-数据库所有者
    '     lngSys-系统号
    '     cnMain-数据库连接
    ' 返回:初始成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-08-13 10:31:00
    '说明:
    '-------------------------------------------------------------------------------------------------
    Set gcnOracle = cnMain
    If Not CheckValid Then Exit Function
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    gstr单位名称 = GetUnitName()
    
    glngSys = lngSys
    gstr单位名称 = GetUnitName()
    gstrDBUser = strDbUser
    InitOracle = True
End Function

Public Function ChargeBillHaveData() As Boolean
     '-------------------------------------------------------------------------------------------------
    '功能:当前轧账是否存在数据
    '返回:存在数据,返回true,否则返回False
    '编制:刘尔旋
    '日期:2017-4-7
    '说明:
    '-------------------------------------------------------------------------------------------------
    On Error GoTo errH
    If mfrmChargeBillTotal.mrsList Is Nothing And mfrmChargeBillTotal.mrsListBill Is Nothing Then
        ChargeBillHaveData = False
        Exit Function
    End If
    If Not mfrmChargeBillTotal.mrsList Is Nothing Then
        If mfrmChargeBillTotal.mrsList.State = 1 Then
            If mfrmChargeBillTotal.mrsList.RecordCount <> 0 Then
                ChargeBillHaveData = True
                Exit Function
            End If
        End If
    End If
    If Not mfrmChargeBillTotal.mrsListBill Is Nothing Then
        If mfrmChargeBillTotal.mrsListBill.State = 1 Then
            If mfrmChargeBillTotal.mrsListBill.RecordCount <> 0 Then
                ChargeBillHaveData = True
                Exit Function
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub ChargeRollingListShow(ByVal frmMain As Object, _
            ByVal bytType As TotalType, ByVal strChargeRollingID As String, ByVal lngModule As Long, _
            ByVal strPrivs As String, Optional ByVal dtStartDate As Date, _
            Optional ByVal dtEndDate As Date, Optional blnDel As Boolean)
    '-------------------------------------------------------------------------------------------------
    '功能:收费员轧账明细显示接口
    '入参:frmMain-调用的主窗体
    '　　bytType:1-收费员轧帐；2-小组收款;3-小组轧帐;4-财务收款。
    '       lngChargeRollingID -收费员的轧帐ID
    '      dtStartDate-可选参数,开始轧帐时间,lngChargeRollIngID=0时，必须传入
    '      dtEndDate-可选参数，结束轧帐时间,lngChargeRollIngID=0时，必须传入
    '      blnDel-是否作废记录
    '编制:刘兴洪
    '日期:2013-08-13 10:31:00
    '说明:
    '-------------------------------------------------------------------------------------------------
    Dim frmNew As frmChargeBillList
    Set frmNew = New frmChargeBillList
    Load frmNew
   Call frmNew.ShowMe(frmMain, lngModule, strPrivs, bytType, strChargeRollingID, dtStartDate, dtEndDate, blnDel)
   If Not frmNew Is Nothing Then Unload frmNew
   Set frmNew = Nothing
 End Sub
 
Public Function ChargeRollingCurtain(ByVal frmMain As Object) As Boolean
        '-------------------------------------------------------------------------------------------------
        '功能:收费员轧账接口
        '入参:frmMain-调用的主窗体
        '        strOperatorName-收费员
        '返回:收费轧帐成功一次以上,返回true,否则返回False
        '编制:刘兴洪
        '日期:2013-08-13 10:31:00
        '说明:
        '-------------------------------------------------------------------------------------------------
        Dim strTitle As String, frmTemp As Form
        glngModul = 1506:
        gstrPrivs = ";" & GetPrivFunc(glngSys, 1506) & ";"
        If zlStr.IsHavePrivs(gstrPrivs, "基本") = False Then
            MsgBox "你没有操作收费轧帐管理的权限,请与系统管理员联系", vbExclamation + vbOKOnly, gstrSysName
            Exit Function
        End If
        strTitle = "收费轧帐管理"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Function
            End If
        Next
        Dim frmNew As frmChargeRollingCurtainMgr
        Set frmNew = New frmChargeRollingCurtainMgr
'        Load frmNew
        ChargeRollingCurtain = frmNew.zlShowChargeRollingCourtain(frmMain, glngModul, gstrPrivs)
'        If Not frmNew Is Nothing Then Unload frmNew
'        Set frmNew = Nothing
End Function

Public Sub CheckCash(ByVal frmMain As Object, Optional ByVal dblMoney As Double)
    '-------------------------------------------------------------------------------------------------
    '功能:清点现金
    '入参:frmMain-调用的主窗体
    '       dblMoney -当前帐面的金额.
    '编制:刘兴洪
    '日期:2013-08-13 10:31:00
    '说明:
    '-------------------------------------------------------------------------------------------------
    frmMoneyEnum.ShowMe frmMain, dblMoney
End Sub

Public Function GetHandIn() As Double
    '-------------------------------------------------------------------------------------------------
    '功能:获取轧帐现金金额
    '编制:刘尔旋
    '日期:2014-12-04 09:34:33
    '说明:
    '-------------------------------------------------------------------------------------------------
    GetHandIn = mfrmChargeBillTotal.GetCashMoney
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省9号字)，1-大号(12号);>1: 为指定的字号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-03 18:05:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call initChargeBillTotalWin
    Call mfrmChargeBillTotal.SetFontSize(bytSize)
End Sub
Public Function GetChargeAndBillTotalForm() As Object
    '------------------------------------------------------------------------
    '功能:获取"收款及票据汇总窗口"对象
    '返回:获取窗体对象
    '编制:刘兴洪
    '日期:2013-08-31 13:26:15
    '说明:由主程序初始化时只调用一次获取,由主窗体绑定该对象
    '-------------------------------------------------------------------------
    Call initChargeBillTotalWin
    Set GetChargeAndBillTotalForm = mfrmChargeBillTotal
 End Function
 Public Sub ClearChargeAndBillTotalForm()
    '------------------------------------------------------------------------
    '功能:清除"收款及票据汇总窗口"对象的所有信息
    '编制:刘兴洪
    '日期:2013-08-31 13:26:15
    '-------------------------------------------------------------------------
    Call initChargeBillTotalWin
    Call mfrmChargeBillTotal.ClearData
 End Sub
 
 Public Function LoadChargeAndBillTotalData(ByVal frmMain As Object, _
      ByVal lngModule As Long, ByVal strPrivs As String, _
      ByVal bytType As TotalType, ByVal lngChargeRollingID As Long, _
      Optional ByVal dtStartDate As Date, Optional ByVal dtEndDate As Date, _
      Optional blnOlnyView As Boolean = True, _
      Optional ByVal blnDel As Boolean = False, _
      Optional strPersonName As String = "", _
      Optional strRollingType As String) As Boolean
    '-------------------------------------------------------------------------------------------------
    '功能:收费员轧账接口
    '入参:frmMain-调用的主窗体
    '    lngModule-模块号
    '    strPrivs-权限串
    '　　bytType:1-收费员轧帐；2-小组收款;3-小组轧帐;
    '            4-财务收款(针对收费员或财务组收款)或财务收款查询;
    '            5-财务收款(仅针对非收费员收款)。
    '    lngChargeRollingID -收费员的轧帐ID
    '    dtStartDate-可选参数,开始轧帐时间,lngChargeRollIngID=0时，必须传入
    '    dtEndDate-可选参数，结束轧帐时间,lngChargeRollIngID=0时，必须传入
    '    blnOlnyView-仅能查看(不能编制票据号码)
    '    blnDel-是否作废记录
    '    strPersonName-指定的收费员,传为"-"时,表示只清除数据
    '    strRollingType-轧帐类别(0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡)
    '返回:数据加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-08-13 10:31:00
    '-------------------------------------------------------------------------------------------------
    Call initChargeBillTotalWin
    LoadChargeAndBillTotalData = mfrmChargeBillTotal.LoadChargeAndBillTotalData(frmMain, lngModule, strPrivs, bytType, lngChargeRollingID, dtStartDate, dtEndDate, blnOlnyView, blnDel, strPersonName, strRollingType)
 End Function
 Public Sub zlPrint(ByVal bytMode As Byte, _
    Optional strDeptName As String = "", Optional strMemo As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出列表信息
    '入参:bytMode=1-打印,2-预览,3-输出到Excel
    '       strDeptName-收款部门名称(收费员轧帐时转入)
    '       strMemo-备注(收费员轧帐时转入)
    '编制:刘兴洪
    '日期:2013-09-13 10:23:30
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmChargeBillTotal.zlPrint(bytMode, strDeptName, strMemo)
End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmChargeBillTotal Is Nothing Then Unload mfrmChargeBillTotal
    Set mfrmChargeBillTotal = Nothing
End Sub



