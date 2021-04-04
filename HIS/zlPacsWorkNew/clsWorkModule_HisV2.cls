VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkModule_HisV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IWorkMenuV2
Implements IWorkModule

Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "HIS"

'HIS模块信息
Private Type THisModuleInfo
    objModule As Object
    strModuleName As String
    lngModuleType As Long

    lngTmpAdviceId As Long
    lngTmpSendNo As Long
    
    blnRefreshState As Boolean
End Type


''HIS模块类型
'Public Enum THisModuleType
'    hmtAll = -1          '所有模块
'    hmtExpense = 0       '费用
'    hmtInAdvice = 1      '住院医嘱
'    hmtOutAdvices = 2    '门诊医嘱
'    hmtInEPRs = 3        '住院病历
'    hmtOutEPRs = 4        '门诊病历
'    hmtEMR = 5           '电子病历
'End Enum

Private mblnNoRegCanPay As Boolean
Private mlngModule  As Long
Private mstrPrivs  As String
Private mlngCurDeptId As Long
'Private mblnIsClinicalDeptPro As Boolean '是否具备临床科室属性
Private mObjNotify As IEventNotify

Private mobjStudyInfo As clsStudyInfo

Private WithEvents mobjInAdvice As zlCISKernel.clsDockInAdvices      '住院医嘱对象
Attribute mobjInAdvice.VB_VarHelpID = -1
Private WithEvents mobjOutAdvice As zlCISKernel.clsDockOutAdvices  '门诊医嘱对象
Attribute mobjOutAdvice.VB_VarHelpID = -1
Private mobjExpense As Object 'zlPublicExpense.clsDockExpense       '费用对象
Private mobjInEPRs As zlRichEPR.cDockInEPRs             '住院病历对象
Private mobjOutEPRs As zlRichEPR.cDockOutEPRs           '门诊病历对象
Private mobjEMR As Object

Private aryRegModule() As THisModuleInfo                       '保存当前创建的模块


Private mObjActiveMenuBar As CommandBars
Private mblnIsHistoryMode As Boolean



'获取菜单接口对象
Property Get zlMenu() As IWorkMenuV2
    Set zlMenu = Me
End Property


'获取模块接口对象
Property Get zlModule() As IWorkModule
    Set zlModule = Me
End Property


Private Sub Class_Initialize()
    '避免使用ubound时出错，因此默认初始化一个空数组元素
    ReDim Preserve aryRegModule(0)
    
    aryRegModule(0) = GetNullModule
End Sub


Public Sub Free()
On Error GoTo errhandle
    Dim i As Long
    
    If SafeArrayGetDim(aryRegModule) <> 0 Then
        For i = LBound(aryRegModule) To UBound(aryRegModule)
            If Not aryRegModule(i).objModule Is Nothing Then
                Unload aryRegModule(i).objModule.zlGetForm
                Set aryRegModule(i).objModule = Nothing
            End If
        Next i
    End If
    
    Set mobjExpense = Nothing
    Set mobjInAdvice = Nothing
    Set mobjOutAdvice = Nothing
    Set mobjInEPRs = Nothing
    Set mobjOutEPRs = Nothing
    Set mobjEMR = Nothing
    
    aryRegModule = Null
    
errhandle:
End Sub


Private Function GetNullModule() As THisModuleInfo
    Set GetNullModule.objModule = Nothing
    
    GetNullModule.strModuleName = ""
    GetNullModule.blnRefreshState = True
    GetNullModule.lngModuleType = hmAll
End Function

Private Sub InitEMR()
    Set mobjEMR = DynamicCreate("zlRichEMR.clsDockEMR", "电子病历")
    If Not mobjEMR Is Nothing Then
        If Not mobjEMR.Init(gobjEmr, gcnOracle, glngSys) Then
          Set mobjEMR = Nothing
        End If
    End If
End Sub
 

Private Sub Class_Terminate()
    Set mObjNotify = Nothing
    Set mobjExpense = Nothing
    Set mobjInEPRs = Nothing
    Set mobjOutEPRs = Nothing
    Set mobjEMR = Nothing
    Set mObjActiveMenuBar = Nothing
    Set mobjInAdvice = Nothing
    Set mobjOutAdvice = Nothing
    Set mobjStudyInfo = Nothing
End Sub

'接口实现部分*********************************************************************************

Public Function IWorkMenuV2_zlBaseMenuID() As Long

End Function

Public Function IWorkMenuV2_zlExecuteCmd(ByVal lngCmdType As Long)
'执行菜单命令

End Function

Public Function IWorkMenuV2_zlIsModuleMenu(ByVal strModuleName As String, objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单
'医嘱，费用，病历模块创建的菜单名对应为“医嘱”，“费用”，“病历”
    IWorkMenuV2_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG Or objControlMenu.Category = "已判断", True, False)
End Function


Public Sub IWorkMenuV2_zlCreateMenu(ByVal strModuleName As String, objMenuBar As Object)
'创建影像记录对应的菜单
    Dim objNewMenu As CommandBarControl
    Dim objHisModule As THisModuleInfo
    Dim lngMenuId As Long

    Set mObjActiveMenuBar = objMenuBar

    Set objNewMenu = objMenuBar.FindControl(, conMenu_EditPopup)

    '如果菜单存在，则先删除对应的菜单和工具栏按钮重新创建
    Call DelModuleMenuAndToolBar(objMenuBar, objNewMenu, M_STR_MODULE_MENU_TAG)

    objHisModule = GetModuleInfo(strModuleName) 'mobjStudyInfo
    
    If objHisModule.objModule Is Nothing Then Exit Sub
    
    Select Case objHisModule.strModuleName
        Case C_WORKMODULE_NAME_住院医嘱, C_WORKMODULE_NAME_门诊医嘱
            Call objHisModule.objModule.zlDefCommandBars(mObjNotify.Owner, mObjActiveMenuBar, 2)

        Case C_WORKMODULE_NAME_费用记录
            Call objHisModule.objModule.zlDefCommandBars(mObjNotify.Owner, mObjActiveMenuBar)
             

        Case C_WORKMODULE_NAME_住院病历, C_WORKMODULE_NAME_门诊病历, C_WORKMODULE_NAME_电子病历
            Call objHisModule.objModule.zlDefCommandBars(mObjActiveMenuBar)
                
    End Select

    Call ConfigMenuCategory(objMenuBar.ActiveMenuBar, M_STR_MODULE_MENU_TAG)
    Call ConfigMenuCategory(objMenuBar.Item(2), M_STR_MODULE_MENU_TAG)

    '避免创建模块工具栏时，窗口左上角出现不完全显示的图标，一闪而过的现象
'    Call objMenuBar.RecalcLayout
End Sub

Private Function GetSubCommandBar(objControl As Object) As Object
On Error GoTo errhandle
    Set GetSubCommandBar = objControl.CommandBar
Exit Function
errhandle:
    Set GetSubCommandBar = Nothing
End Function


Private Sub ConfigMenuCategory(objCommandBar As Object, ByVal strCategory As String)
'配置Category属性
    Dim objSubMenu As CommandBar
    
    Dim i As Long
    Dim j As Long
    
    If objCommandBar Is Nothing Then Exit Sub

    
    For i = 1 To objCommandBar.Controls.Count
        If objCommandBar.Controls(i).Category = "" Then objCommandBar.Controls(i).Category = strCategory
        
        Set objSubMenu = GetSubCommandBar(objCommandBar.Controls(i))
        If Not objSubMenu Is Nothing Then Call ConfigMenuCategory(objSubMenu, strCategory)
    Next i
End Sub


Private Sub DelModuleMenuAndToolBar(objMenuBar As Object, objModuleMenu As CommandBarControl, ByVal strMenuName As String)
'删除模块菜单和工具栏
    Dim objToolBar As CommandBar
    Dim objActiveMenu As CommandBar
    Dim objSubMenu As CommandBar
    
    Dim i As Long
    Dim j As Long
    
    If Not objModuleMenu Is Nothing Then
        Call objModuleMenu.Delete
    End If
    
    '删除在其他菜单下创建的模块菜单
    Set objActiveMenu = objMenuBar.ActiveMenuBar
    For i = objActiveMenu.Controls.Count To 1 Step -1
        If objActiveMenu.Controls(i).type <> xtpControlButton Then
            Set objSubMenu = objActiveMenu.Controls(i).CommandBar
            
            For j = objSubMenu.Controls.Count To 1 Step -1
                If objSubMenu.Controls(j).Category = strMenuName Or objSubMenu.Controls(j).Category = "已判断" Then
                    Call objSubMenu.Controls(j).Delete
                End If
            Next j
            
            If objActiveMenu.Controls(i).Category = strMenuName Or objActiveMenu.Controls(i).Category = "已判断" Then
                Call objActiveMenu.Controls(i).Delete
            End If
        End If
    Next i
    
    '删除工具栏中的模块对应按钮
    Set objToolBar = objMenuBar.Item(2)
    If Not objToolBar Is Nothing Then
        For i = objToolBar.Controls.Count To 1 Step -1
            If objToolBar.Controls(i).Category = strMenuName Or objToolBar.Controls(i).Category = "已判断" Then
                Call objToolBar.Controls(i).Delete
            Else
                If objToolBar.Controls(i).Category = "" Then
                    objToolBar.Controls(i).Category = "TEMP"
                End If
            End If
        Next i
    End If
End Sub


Private Function GetNewMenu(objMenuBar As Object, ByVal strMenuName As String) As CommandBarControl
'获取新创建的菜单
    Dim i As Long
    Dim objActiveMenu As CommandBar
    
    Set objActiveMenu = objMenuBar.ActiveMenuBar
    
    For i = 1 To objActiveMenu.Controls.Count
        If InStr(objActiveMenu.Controls(i).Caption, strMenuName) > 0 Then
            Set GetNewMenu = objActiveMenu.Controls(i)
            Exit Function
        End If
    Next i
    
    Set GetNewMenu = Nothing
End Function


Public Sub IWorkMenuV2_zlCreateToolBar(ByVal strModuleName As String, objToolBar As Object)
'创建工具栏
    Exit Sub
End Sub


Public Sub IWorkMenuV2_zlClearMenu(ByVal strModuleName As String)
'清除所创建的菜单
    Dim objNewMenu As CommandBarControl
    
    If mObjActiveMenuBar Is Nothing Then Exit Sub
    
    Set objNewMenu = mObjActiveMenuBar.FindControl(, conMenu_EditPopup)

    '如果菜单存在，则先删除对应的菜单和工具栏按钮重新创建
    Call DelModuleMenuAndToolBar(mObjActiveMenuBar, objNewMenu, M_STR_MODULE_MENU_TAG)
End Sub


Public Sub IWorkMenuV2_zlClearToolBar(ByVal strModuleName As String)
'清除创建的工具栏
    Exit Sub
End Sub

Public Sub IWorkMenuV2_zlExecuteMenu(ByVal strModuleName As String, ByVal lngMenuId As Long)
'根据菜单ID执行对应功能
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    Dim objHisModule As THisModuleInfo
 

    objHisModule = GetModuleInfo(strModuleName) ', mobjStudyInfo
    If objHisModule.objModule Is Nothing Then Exit Sub
    
    
    Set objCbrControl = mObjActiveMenuBar.FindControl(, lngMenuId, , True)
    If objCbrControl Is Nothing Then Exit Sub
    
    
    Select Case objHisModule.strModuleName
        Case C_WORKMODULE_NAME_住院医嘱, C_WORKMODULE_NAME_门诊医嘱
            Call objHisModule.objModule.zlExecuteCommandBars(objCbrControl)

        Case C_WORKMODULE_NAME_费用记录
            Call objHisModule.objModule.zlExecuteCommandBars(objCbrControl)

        Case C_WORKMODULE_NAME_住院病历, C_WORKMODULE_NAME_门诊病历, C_WORKMODULE_NAME_电子病历
            Call objHisModule.objModule.zlExecuteCommandBars(objCbrControl)
                
    End Select
End Sub


Public Sub IWorkMenuV2_zlUpdateMenu(ByVal strModuleName As String, Control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    Dim objHisModule As THisModuleInfo
    Dim lngMenuId As Long
 
    If Control Is Nothing Then Exit Sub
    
    Control.Enabled = True
    
    If mobjStudyInfo Is Nothing Then
        Control.Enabled = False
        Exit Sub
    End If
    
    objHisModule = GetModuleInfo(strModuleName) ', mobjStudyInfo
    If objHisModule.objModule Is Nothing Then Exit Sub
    
    
    Select Case objHisModule.strModuleName
        Case C_WORKMODULE_NAME_住院医嘱, C_WORKMODULE_NAME_门诊医嘱
            Call objHisModule.objModule.zlUpdateCommandBars(Control)

        Case C_WORKMODULE_NAME_费用记录
            Call objHisModule.objModule.zlUpdateCommandBars(Control)
                If Control.ID = conMenu_Edit_NewItem Then
                    If InStr(Control.Caption, "补附费") > 0 Or InStr(Control.Caption, "补充附加费用") > 0 Then
                        Control.Enabled = Control.Enabled And (mobjStudyInfo.intStep >= 2 Or mblnNoRegCanPay) And Not mblnIsHistoryMode
                    End If
                End If
             
        Case C_WORKMODULE_NAME_住院病历, C_WORKMODULE_NAME_门诊病历, C_WORKMODULE_NAME_电子病历
            Call objHisModule.objModule.zlUpdateCommandBars(Control)
                
    End Select
    
End Sub


Public Sub IWorkMenuV2_zlPopupMenu(ByVal strModuleName As String, objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单
    Exit Sub
End Sub

Public Sub IWorkMenuV2_zlRefreshSubMenu(ByVal strModuleName As String, objCommandBar As Object)
'刷新弹出的子菜单
    Dim objHisModule As THisModuleInfo

    objHisModule = GetModuleInfo(strModuleName)
    If objHisModule.objModule Is Nothing Then Exit Sub
 
    Call objHisModule.objModule.zlPopupCommandBars(objCommandBar)

    Call ConfigMenuCategory(objCommandBar, M_STR_MODULE_MENU_TAG)
    Call ConfigMenuCategory(mObjActiveMenuBar.Item(2), M_STR_MODULE_MENU_TAG)
End Sub
'*********************************************************************************************


Private Function HasClinicalPro(ByVal lngDepartmentId As Long) As Long
'判断当前科室是否具备临床属性
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As ADODB.Recordset

    HasClinicalPro = False
    
    strSQL = "select 部门ID from 部门性质说明 where 部门ID=[1] and 工作性质='临床'"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "判断部门性质", lngDepartmentId)
    
    If rsData.RecordCount > 0 Then
        '不具备临床性质，则查询医生默认所在科室ID
        HasClinicalPro = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetModuleInfo(ByVal strModuleName As String) As THisModuleInfo
    Dim i As Long
    Dim strInsideModule As String
    
'    strInsideModule = ConvertInsideModule(strModuleName, objStudyInfo)
    
    GetModuleInfo = GetNullModule
    For i = 1 To UBound(aryRegModule)
        If aryRegModule(i).strModuleName = strModuleName Then
            GetModuleInfo = aryRegModule(i)
            Exit Function
        End If
    Next
End Function


Public Function GetInsideHisModule(ByVal strModuleName As String, ByVal lngPatientFrom As Long) As String
'获取HIS模块类型
'lngPatientFrom：病人来源 2-住院，其余表示外来
    Select Case strModuleName
        Case C_TAB_NAME_医嘱记录
            If lngPatientFrom <> 2 Then
                GetInsideHisModule = "门诊医嘱"
            Else
                GetInsideHisModule = "住院医嘱"
            End If
            
        Case C_TAB_NAME_病历记录
            If lngPatientFrom <> 2 Then
                GetInsideHisModule = "门诊病历"
            Else
                GetInsideHisModule = "住院病历"
            End If
            
        Case C_TAB_NAME_电子病历
            GetInsideHisModule = "电子病历"
            
        Case C_TAB_NAME_费用记录
            GetInsideHisModule = "费用记录"
    End Select
End Function

Property Get IWorkModule_AdviceId() As Long
    IWorkModule_AdviceId = mobjStudyInfo.lngAdviceId
End Property

Property Get IWorkModule_DeptId() As Long
    IWorkModule_DeptId = mlngCurDeptId
End Property

Public Sub IWorkModule_zlInitModule(objNotify As IEventNotify, ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDeptId As String)
'初始化模块参数
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngCurDeptId = lngDeptId
    mblnIsHistoryMode = False
    
    '判断执行科室是否具备临床性质
    'mblnIsClinicalDeptPro = HasClinicalPro(lngDepartId)
    mblnNoRegCanPay = GetDeptPara(lngDeptId, "允许未报到补费", 0) = "1" '       '无报告或报告未签名允许完成
    
    Set mObjNotify = objNotify
    
    If Not mobjExpense Is Nothing And mlngModule = G_LNG_PATHOLSYS_NUM Then
        Call mobjExpense.zlInitModule(lngModule, strPrivs, lngDeptId, mObjNotify.Owner)
    End If
'    Call gobjKernel.InitCISKernel(gcnOracle, gfrmMain, glngSys, strPrivs) '初始化医嘱，病历核心部件
'    Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
End Sub


Public Function IWorkModule_zlGetModule(ByVal strSelModuleName As String, Optional ByVal blnAutoCreate As Boolean = True) As Object
'获取指定模块对象
    Dim objTmp As zlPublicExpense.clsPublicExpense
    Dim strInsideModule As String
 
    
    Set IWorkModule_zlGetModule = Nothing

'    strInsideModule = ConvertInsideModule(strSelModuleName, objStudyInfo)
    
    Select Case strSelModuleName
        Case C_WORKMODULE_NAME_费用记录  '费用
            If mobjExpense Is Nothing And blnAutoCreate Then
                If mlngModule = G_LNG_PATHOLSYS_NUM Then
                    Set mobjExpense = New frmPatholPrice

                    zlControl.FormSetCaption mobjExpense, False, False
                Else
                    Set objTmp = New zlPublicExpense.clsPublicExpense
                    Call objTmp.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
                    Set mobjExpense = objTmp.zlDockExpense
                End If

                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjExpense

                aryRegModule(UBound(aryRegModule)).strModuleName = C_WORKMODULE_NAME_费用记录
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmExpense
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If

            Set IWorkModule_zlGetModule = mobjExpense.zlGetForm()

        Case C_WORKMODULE_NAME_住院医嘱 '住院医嘱
            If mobjInAdvice Is Nothing And blnAutoCreate Then
                Set mobjInAdvice = New zlCISKernel.clsDockInAdvices

                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjInAdvice

                aryRegModule(UBound(aryRegModule)).strModuleName = C_WORKMODULE_NAME_住院医嘱
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmInAdvice
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If

            Set IWorkModule_zlGetModule = mobjInAdvice.zlGetForm()

        Case C_WORKMODULE_NAME_门诊医嘱  '门诊医嘱
            If mobjOutAdvice Is Nothing And blnAutoCreate Then
                Set mobjOutAdvice = New zlCISKernel.clsDockOutAdvices

                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjOutAdvice

                aryRegModule(UBound(aryRegModule)).strModuleName = C_WORKMODULE_NAME_门诊医嘱
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmOutAdvices
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If

            Set IWorkModule_zlGetModule = mobjOutAdvice.zlGetForm()

        Case C_WORKMODULE_NAME_住院病历 '住院病历
            If mobjInEPRs Is Nothing And blnAutoCreate Then
                Set mobjInEPRs = New zlRichEPR.cDockInEPRs

                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjInEPRs

                aryRegModule(UBound(aryRegModule)).strModuleName = C_WORKMODULE_NAME_住院病历
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmInEPRs
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If

            Set IWorkModule_zlGetModule = mobjInEPRs.zlGetForm()

        Case C_WORKMODULE_NAME_门诊病历 '门诊病历
            If mobjOutEPRs Is Nothing And blnAutoCreate Then
                Set mobjOutEPRs = New zlRichEPR.cDockOutEPRs

                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjOutEPRs

                aryRegModule(UBound(aryRegModule)).strModuleName = C_WORKMODULE_NAME_门诊病历
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmOutAdvices
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If

            Set IWorkModule_zlGetModule = mobjOutEPRs.zlGetForm()

        Case C_WORKMODULE_NAME_电子病历 '电子病历
            If mobjEMR Is Nothing And blnAutoCreate Then
                Call InitEMR

                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjEMR

                aryRegModule(UBound(aryRegModule)).strModuleName = C_WORKMODULE_NAME_电子病历
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmEMR
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If

            Set IWorkModule_zlGetModule = mobjEMR.zlGetForm()
    End Select
End Function


Public Sub IWorkModule_zlRefresh(objStudyInfo As clsStudyInfo, ByVal strModuleName As String, _
    Optional ByVal blnIsForceRefresh As Boolean = False, Optional ByVal blnIsHistory As Boolean = False)
'刷新界面
    Dim rsTemp As ADODB.Recordset
    Dim iRows As Integer
    Dim iCols As Integer
    Dim strStudyUID As String
    Dim objHisModule As THisModuleInfo
    Dim strSQL As String
    Dim lngBillID As Long
    
    On Error GoTo errhandle
    
    mblnIsHistoryMode = blnIsHistory
     
    If Not mobjStudyInfo Is Nothing And Not objStudyInfo Is Nothing Then
        If mobjStudyInfo.IsEquals(objStudyInfo) And blnIsForceRefresh = False Then Exit Sub
    End If
    
    Set mobjStudyInfo = objStudyInfo
    
    objHisModule = GetModuleInfo(strModuleName) ', mobjStudyInfo
    
    If objHisModule.objModule Is Nothing Then Exit Sub
    
    objHisModule.lngTmpAdviceId = mobjStudyInfo.lngAdviceId
    objHisModule.lngTmpSendNo = mobjStudyInfo.lngSendNo
    
    objHisModule.blnRefreshState = True
    
    Select Case objHisModule.strModuleName
        Case C_WORKMODULE_NAME_费用记录
            If mlngModule = G_LNG_PATHOLSYS_NUM Then
                objHisModule.objModule.zlRefresh mlngCurDeptId, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.blnMoved
            Else
                objHisModule.objModule.zlRefresh mlngCurDeptId, mobjStudyInfo.lngAdviceId & ":" & mobjStudyInfo.lngSendNo & ":0", mobjStudyInfo.blnMoved
            End If
        Case C_WORKMODULE_NAME_住院医嘱
            objHisModule.objModule.zlRefresh mobjStudyInfo.lngPatId, mobjStudyInfo.lngPageID, mobjStudyInfo.lngUnit, mobjStudyInfo.lngPatDept, _
                                        IIf(blnIsHistory, -999, mobjStudyInfo.intState), mobjStudyInfo.blnMoved, mobjStudyInfo.lngAdviceId, mobjStudyInfo.intState, mlngCurDeptId
            
        Case C_WORKMODULE_NAME_门诊医嘱
            If mobjStudyInfo.strRegNo = "" Then   '自行登记的病人没有挂号单号
                objHisModule.objModule.zlRefresh mobjStudyInfo.lngPatId, "", False, , , mlngCurDeptId
            Else
                objHisModule.objModule.zlRefresh mobjStudyInfo.lngPatId, mobjStudyInfo.strRegNo, _
                                            IIf(blnIsHistory, False, mobjStudyInfo.blnIsInsidePatient), mobjStudyInfo.blnMoved, mobjStudyInfo.lngAdviceId, mlngCurDeptId
            End If
            
        Case C_WORKMODULE_NAME_住院病历
            objHisModule.objModule.zlRefresh mobjStudyInfo.lngPatId, mobjStudyInfo.lngPageID, mlngCurDeptId, _
                                        IIf(blnIsHistory, False, True), mobjStudyInfo.blnMoved, 0, True
            
        Case C_WORKMODULE_NAME_门诊病历
            strSQL = "Select Id From 病人挂号记录 Where No = [1]"
            If mobjStudyInfo.blnMoved Then
                strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
            End If
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询挂号ID", mobjStudyInfo.strRegNo)
            If rsTemp.EOF = False Then
                lngBillID = nvl(rsTemp!ID)
            End If
            
            objHisModule.objModule.zlRefresh mobjStudyInfo.lngPatId, lngBillID, mlngCurDeptId, _
                                        IIf(blnIsHistory, False, True), mobjStudyInfo.blnMoved, True, 0
        
        Case C_WORKMODULE_NAME_电子病历
            If mobjStudyInfo.lngPatientFrom = 1 Then
                strSQL = "Select Id From 病人挂号记录 Where No = [1]"
                If mobjStudyInfo.blnMoved Then
                    strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
                End If
            
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询挂号ID", mobjStudyInfo.strRegNo)
                If rsTemp.EOF = False Then
                    lngBillID = nvl(rsTemp!ID)
                End If
            End If
            
            objHisModule.objModule.zlRefresh mobjStudyInfo.lngPatId, IIf(mobjStudyInfo.lngPatientFrom = 2, mobjStudyInfo.lngPageID, lngBillID), mlngCurDeptId, IIf(mobjStudyInfo.lngPatientFrom = 2, 0, 2), IIf(mobjStudyInfo.lngPatientFrom = 2, 2, 1)
    End Select
    
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mobjInAdvice_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'打开报告
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    If mObjNotify Is Nothing Then Exit Sub

    strSQL = "Select 医嘱ID, 0 as Moved from 病人医嘱报告 where 病历ID=[1] " & vbCrLf & _
                    " Union " & vbCrLf & _
                    "Select 医嘱ID, 1 as Moved from H病人医嘱报告 where 病历ID=[1] "
                    
                    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告医嘱信息", 报告ID)
    If rsData.RecordCount <= 0 Then Exit Sub

    If CanPrint Then
        Call mObjNotify.SendRequest(WM_REPORT_PRINT, , Val(nvl(rsData!医嘱ID)), Val(nvl(rsData!Moved)), 报告ID)
    Else
        Call mObjNotify.SendRequest(WM_REPORT_VIEW, , Val(nvl(rsData!医嘱ID)), Val(nvl(rsData!Moved)), 报告ID)
    End If
End Sub

Private Sub mobjOutAdvice_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'打开报告
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    If mObjNotify Is Nothing Then Exit Sub

    strSQL = "Select 医嘱ID, 0 as Moved from 病人医嘱报告 where 病历ID=[1] " & vbCrLf & _
                    " Union " & vbCrLf & _
                    "Select 医嘱ID, 1 as Moved from H病人医嘱报告 where 病历ID=[1] "
                    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告医嘱信息", 报告ID)
    If rsData.RecordCount <= 0 Then Exit Sub

    If CanPrint Then
        Call mObjNotify.SendRequest(WM_REPORT_PRINT, , Val(nvl(rsData!医嘱ID)), Val(nvl(rsData!Moved)), 报告ID)
    Else
        Call mObjNotify.SendRequest(WM_REPORT_VIEW, , Val(nvl(rsData!医嘱ID)), Val(nvl(rsData!Moved)), 报告ID)
    End If
End Sub

Private Sub mobjInAdvice_ViewPACSImage(ByVal 医嘱ID As Long)
'打开观片
    If mObjNotify Is Nothing Then Exit Sub
    Call mObjNotify.SendRequest(WM_IMG_OPENVIEW, , 医嘱ID)
End Sub

Private Sub mobjOutAdvice_ViewPACSImage(ByVal 医嘱ID As Long)
'打开观片
    If mObjNotify Is Nothing Then Exit Sub
    Call mObjNotify.SendRequest(WM_IMG_OPENVIEW, , 医嘱ID)
End Sub



