VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkModule_His"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IWorkMenu


Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "HIS"

'HIS模块信息
Public Type TModule
    objModule As Object
    strModuleName As String
    lngModuleType As Long

    lngTmpAdviceId As Long
    lngTmpSendNo As Long
    
    blnRefreshState As Boolean
End Type


'HIS模块类型
Public Enum THisModule
    hmAll = -1          '所有模块
    hmExpense = 0       '费用
    hmInAdvice = 1      '住院医嘱
    hmOutAdvices = 2    '门诊医嘱
    hmInEPRs = 3        '住院病历
    hmOutEPRs = 4        '门诊病历
    hmEMR = 5           '电子病历
End Enum


Private mlngModule  As Long
Private mstrPrivs  As String
Private mlngCurDeptId As Long
'Private mblnIsClinicalDeptPro As Boolean '是否具备临床科室属性
Private mobjOwner As Object

Private mlngAdviceID As Long
Private mlngSendNo As Long
Private mblnMoved As Boolean
Private mlngStudyState As Long

Private mblnReadOnly As Boolean

'各个模块所需的一些信息

Private mlngPatientId As Long
Private mlngUnitId As Long
Private mlngPatientDepartId As Long
Private mintState As Long
Private mstrRegNo As String
Private mblnIsHistory As Boolean
Private mlngPageID As Long
Private mblnIsInsidePatient As Boolean

Private WithEvents mobjInAdvice As zlCISKernel.clsDockInAdvices    '住院医嘱对象
Attribute mobjInAdvice.VB_VarHelpID = -1
Private WithEvents mobjOutAdvice As zlCISKernel.clsDockOutAdvices  '门诊医嘱对象
Attribute mobjOutAdvice.VB_VarHelpID = -1
Private mobjExpense As Object 'zlPublicExpense.clsDockExpense       '费用对象
Private mobjInEPRs As zlRichEPR.cDockInEPRs             '住院病历对象
Private mobjOutEPRs As zlRichEPR.cDockOutEPRs           '门诊病历对象
Private mobjEMR As Object

Private aryRegModule() As TModule                       '保存当前创建的模块


Private mObjActiveMenuBar As CommandBars



'获取菜单接口对象
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property



'Public Function GetExpenseObj() As Object
'    Set GetExpenseObj = mobjExpense
'End Function
Public Sub NotificationRefresh(Optional ByVal lngModule As Long = -1)
'通知刷新
    Dim i As Long
    
    For i = LBound(aryRegModule) To UBound(aryRegModule)
        If lngModule = -1 Then
            aryRegModule(i).blnRefreshState = False
        Else
            If aryRegModule(i).lngModuleType = lngModule Then
                aryRegModule(i).blnRefreshState = False
                Exit For
            End If
        End If
    Next i
End Sub

Private Sub Class_Initialize()
    '避免使用ubound时出错，因此默认初始化一个空数组元素
    ReDim Preserve aryRegModule(1)
    
    aryRegModule(0) = GetNullModule
End Sub


Public Sub Free()
On Error GoTo ErrHandle
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
    
ErrHandle:
End Sub


Private Function GetNullModule() As TModule
    Set GetNullModule.objModule = Nothing
    
    GetNullModule.strModuleName = ""
    GetNullModule.blnRefreshState = True
    GetNullModule.lngModuleType = hmAll
End Function


Public Function GetExpenseObj() As Object
     If Not mobjExpense Is Nothing Then Set GetExpenseObj = mobjExpense
End Function

Public Function GetAdviceObj() As Object

    If Not mobjInAdvice Is Nothing Or Not mobjOutAdvice Is Nothing Then
    
        If Not mobjInAdvice Is Nothing Then
            Set GetAdviceObj = IIf(mobjInAdvice.zlGetForm.Visible = True, mobjInAdvice, mobjOutAdvice)
        Else
            Set GetAdviceObj = IIf(mobjOutAdvice.zlGetForm.Visible = True, mobjOutAdvice, mobjInAdvice)
        End If
        
    Else
        Set GetAdviceObj = Nothing
    End If

End Function

Public Function GetEMRsObj() As Object
    Set GetEMRsObj = mobjEMR
End Function

Public Function GetEPRsObj() As Object

    If Not mobjInEPRs Is Nothing Or Not mobjOutEPRs Is Nothing Then
        
        If Not mobjInEPRs Is Nothing Then
            Set GetEPRsObj = IIf(mobjInEPRs.zlGetForm.Visible = True, mobjInEPRs, mobjOutEPRs)
        Else
            Set GetEPRsObj = IIf(mobjOutEPRs.zlGetForm.Visible = True, mobjOutEPRs, mobjInEPRs)
        End If
        
    Else
        Set GetEPRsObj = Nothing
    End If
    
End Function

Public Function GetModule(ByVal lngModule As Long, Optional ByVal blnAutoCreate As Boolean = True) As Object
'获取指定模块对象
    Dim objTmp As zlPublicExpense.clsPublicExpense
    Set GetModule = Nothing
    
    Select Case lngModule
        Case hmExpense  '费用
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
                
                aryRegModule(UBound(aryRegModule)).strModuleName = "申请费用"
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmExpense
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If
            
            Set GetModule = mobjExpense.zlGetForm()
            
        Case hmInAdvice '住院医嘱
            If mobjInAdvice Is Nothing And blnAutoCreate Then
                Set mobjInAdvice = New zlCISKernel.clsDockInAdvices
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjInAdvice
                
                aryRegModule(UBound(aryRegModule)).strModuleName = "住院医嘱"
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmInAdvice
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If
            
            Set GetModule = mobjInAdvice.zlGetForm()
               
        Case hmOutAdvices '门诊医嘱
            If mobjOutAdvice Is Nothing And blnAutoCreate Then
                Set mobjOutAdvice = New zlCISKernel.clsDockOutAdvices
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjOutAdvice
                
                aryRegModule(UBound(aryRegModule)).strModuleName = "门诊医嘱"
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmOutAdvices
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If
            
            Set GetModule = mobjOutAdvice.zlGetForm()
            
        Case hmInEPRs '住院病历
            If mobjInEPRs Is Nothing And blnAutoCreate Then
                Set mobjInEPRs = New zlRichEPR.cDockInEPRs
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjInEPRs
                
                aryRegModule(UBound(aryRegModule)).strModuleName = "住院病历"
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmInEPRs
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If
            
            Set GetModule = mobjInEPRs.zlGetForm()
            
        Case hmOutEPRs '门诊病历
            If mobjOutEPRs Is Nothing And blnAutoCreate Then
                Set mobjOutEPRs = New zlRichEPR.cDockOutEPRs
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjOutEPRs
                
                aryRegModule(UBound(aryRegModule)).strModuleName = "门诊病历"
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmOutAdvices
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If
            
            Set GetModule = mobjOutEPRs.zlGetForm()
        
        Case hmEMR '电子病历
            If mobjEMR Is Nothing And blnAutoCreate Then
                Call InitEMR
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)).objModule = mobjEMR
                
                aryRegModule(UBound(aryRegModule)).strModuleName = "电子病历"
                aryRegModule(UBound(aryRegModule)).lngModuleType = hmEMR
                aryRegModule(UBound(aryRegModule)).blnRefreshState = True
            End If
            
            Set GetModule = mobjEMR.zlGetForm()
    End Select
End Function

Private Sub InitEMR()
    Set mobjEMR = DynamicCreate("zlRichEMR.clsDockEMR", "电子病历")
    If Not mobjEMR Is Nothing Then
        If Not mobjEMR.Init(gobjEmr, gcnOracle, glngSys) Then
          Set mobjEMR = Nothing
        End If
    End If
End Sub

Public Function GetActiveModule() As TModule
'获取当前显示得工作模块
    Dim i As Long
    Dim objCurModule As Object
    
    GetActiveModule = GetNullModule
    
    For i = LBound(aryRegModule) To UBound(aryRegModule)
        If Not aryRegModule(i).objModule Is Nothing Then
            Set objCurModule = aryRegModule(i).objModule.zlGetForm()
            
            If objCurModule.Visible Then
                GetActiveModule = aryRegModule(i)
                Exit For
            End If
        End If
    Next i
End Function

'接口实现部分*********************************************************************************

Public Function IWorkMenu_zlIsModuleMenu(ByVal objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单
'医嘱，费用，病历模块创建的菜单名对应为“医嘱”，“费用”，“病历”
    
    IWorkMenu_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG Or objControlMenu.Category = "已判断", True, False)
End Function


Public Sub IWorkMenu_zlCreateMenu(objMenuBar As Object)
'创建影像记录对应的菜单
    Dim objHisModule As TModule
    Dim objNewMenu As CommandBarControl

    Set mObjActiveMenuBar = objMenuBar
    
    
    Set objNewMenu = objMenuBar.FindControl(, conMenu_EditPopup)
    
    '如果菜单存在，则先删除对应的菜单和工具栏按钮重新创建
    Call DelModuleMenuAndToolBar(objMenuBar, objNewMenu, M_STR_MODULE_MENU_TAG)
        
    objHisModule = GetActiveModule()
    
    If Not objHisModule.objModule Is Nothing Then
        Select Case objHisModule.strModuleName
            Case "住院医嘱", "门诊医嘱"
                Call objHisModule.objModule.zlDefCommandBars(mobjOwner, mObjActiveMenuBar, 2)
                
                Set objNewMenu = GetNewMenu(objMenuBar, "医嘱")
                
            Case "申请费用"
                Call objHisModule.objModule.zlDefCommandBars(mobjOwner, mObjActiveMenuBar)
                
                Set objNewMenu = GetNewMenu(objMenuBar, "费用")
                
            Case "住院病历", "门诊病历", "电子病历"
                Call objHisModule.objModule.zlDefCommandBars(mObjActiveMenuBar)
                
                Set objNewMenu = GetNewMenu(objMenuBar, "病历")
                
        End Select
        
        Call ConfigMenuCategory(objMenuBar.ActiveMenuBar, M_STR_MODULE_MENU_TAG)
        Call ConfigMenuCategory(objMenuBar.Item(2), M_STR_MODULE_MENU_TAG)
        
        '避免创建模块工具栏时，窗口左上角出现不完全显示的图标，一闪而过的现象
        Call objMenuBar.RecalcLayout
    End If
End Sub

Private Function GetSubCommandBar(objControl As Object) As Object
On Error GoTo ErrHandle
    Set GetSubCommandBar = objControl.CommandBar
Exit Function
ErrHandle:
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


Public Sub IWorkMenu_zlCreateToolBar(objToolBar As Object)
'创建工具栏
    Exit Sub
End Sub


Public Sub IWorkMenu_zlClearMenu()
'清除所创建的菜单
    Exit Sub
End Sub


Public Sub IWorkMenu_zlClearToolBar()
'清除创建的工具栏
    Exit Sub
End Sub

Public Sub IWorkMenu_zlExecuteMenu(ByVal lngMenuId As Long)
'根据菜单ID执行对应功能
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    Dim objHisModule As TModule
    
    objHisModule = GetActiveModule()
    
    If Not objHisModule.objModule Is Nothing Then
        Set objCbrControl = mObjActiveMenuBar.FindControl(, lngMenuId, , True)
            
        If Not objCbrControl Is Nothing Then Call objHisModule.objModule.zlExecuteCommandBars(objCbrControl)
    End If
    
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim objHisModule As TModule
    
    objHisModule = GetActiveModule()
    
    If Not objHisModule.objModule Is Nothing Then
        Call objHisModule.objModule.zlUpdateCommandBars(control)
    End If
End Sub


Public Sub IWorkMenu_zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单
    Exit Sub
End Sub

Public Sub IWorkMenu_zlRefreshSubMenu(objCommandBar As Object)
'刷新弹出的子菜单
    Dim objHisModule As TModule
    
    objHisModule = GetActiveModule()
    
    If Not objHisModule.objModule Is Nothing Then
        Call objHisModule.objModule.zlPopupCommandBars(objCommandBar)
    End If
    
    Call ConfigMenuCategory(objCommandBar, M_STR_MODULE_MENU_TAG)
    Call ConfigMenuCategory(mObjActiveMenuBar.Item(2), M_STR_MODULE_MENU_TAG)
End Sub
'*********************************************************************************************

Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing)
'初始化模块参数
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngCurDeptId = lngDepartId
    
    '判断执行科室是否具备临床性质
    'mblnIsClinicalDeptPro = HasClinicalPro(lngDepartId)
    
    If Not owner Is Nothing Then Set mobjOwner = owner
    
    If Not mobjExpense Is Nothing And mlngModule = G_LNG_PATHOLSYS_NUM Then
        Call mobjExpense.zlInitModule(lngModule, strPrivs, lngDepartId, owner)
    End If
'    Call gobjKernel.InitCISKernel(gcnOracle, gfrmMain, glngSys, strPrivs) '初始化医嘱，病历核心部件
'    Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
End Sub


Private Function HasClinicalPro(ByVal lngDepartmentId As Long) As Long
'判断当前科室是否具备临床属性
    Dim strSQL As String
    Dim rsData As ADODB.Recordset

    HasClinicalPro = False
    
    strSQL = "select 部门ID from 部门性质说明 where 部门ID=[1] and 工作性质='临床'"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "判断部门性质", lngDepartmentId)
    
    If rsData.RecordCount > 0 Then
        '不具备临床性质，则查询医生默认所在科室ID
        HasClinicalPro = True
    End If
    
End Function


Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, _
    ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'更新医嘱信息
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    mblnReadOnly = False
    

    '如果为完成，登记，拒绝，审核，则不能对检查进行处理
    If lngStudyState = 6 Or lngStudyState = 5 Or lngStudyState = 0 Or lngStudyState = 1 Or lngStudyState = -2 Then
        mblnReadOnly = True
    End If
End Sub

Public Sub zlUpdateOtherInf(ByVal lngPatientID As Long, ByVal lngUnitId As Long, ByVal lngPatientDepartId As Long, _
    ByVal lngPageID As Long, ByVal intState As Long, ByVal strRegNo As String, ByVal blnIsHistory As Boolean, ByVal blnIsInsidePatient As Boolean)
'更新医嘱相关的其他信息
    mlngPatientId = lngPatientID
    mlngUnitId = lngUnitId
    mlngPatientDepartId = lngPatientDepartId
    mintState = intState
    mstrRegNo = strRegNo
    mblnIsHistory = blnIsHistory
    mlngPageID = lngPageID
    mblnIsInsidePatient = blnIsInsidePatient
End Sub


Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False, Optional ByVal lngPatientFrom As Long = 3)
'刷新界面
    Dim rsTemp As ADODB.Recordset
    Dim iRows As Integer
    Dim iCols As Integer
    Dim strStudyUID As String
    Dim objHisModule As TModule
    Dim strSQL As String
    Dim lngBillID As Long
    
    On Error GoTo ErrHandle
    
    objHisModule = GetActiveModule
    
    If objHisModule.objModule Is Nothing Then Exit Sub
    
    '当医嘱改变时，或第一次加载界面时，或强制刷新时才进行界面数据刷新操作
    If (objHisModule.lngTmpAdviceId = mlngAdviceID And objHisModule.lngTmpSendNo = mlngSendNo _
        And objHisModule.blnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    objHisModule.lngTmpAdviceId = mlngAdviceID
    objHisModule.lngTmpSendNo = mlngSendNo
    
    objHisModule.blnRefreshState = True
    
    Select Case objHisModule.strModuleName
        Case "申请费用"
            If mlngModule = G_LNG_PATHOLSYS_NUM Then
                objHisModule.objModule.zlRefresh mlngCurDeptId, mlngAdviceID, mlngSendNo, mblnMoved
            Else
                objHisModule.objModule.zlRefresh mlngCurDeptId, mlngAdviceID & ":" & mlngSendNo & ":0"
            End If
        Case "住院医嘱"
            objHisModule.objModule.zlRefresh mlngPatientId, mlngPageID, mlngUnitId, mlngPatientDepartId, 0, mblnMoved, mlngAdviceID, mintState, mlngCurDeptId
            
        Case "门诊医嘱"
            If mstrRegNo = "" Then   '自行登记的病人没有挂号单号
                objHisModule.objModule.zlRefresh mlngPatientId, "", False, , , mlngCurDeptId
            Else
                objHisModule.objModule.zlRefresh mlngPatientId, mstrRegNo, Not mblnIsHistory And mblnIsInsidePatient, mblnMoved, mlngAdviceID, mlngCurDeptId
            End If
            
        Case "住院病历"
            objHisModule.objModule.zlRefresh mlngPatientId, mlngPageID, mlngCurDeptId, Not mblnIsHistory, mblnMoved, 0, True
            
        Case "门诊病历"
            strSQL = "Select Id From 病人挂号记录 Where No = [1]"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询挂号ID", mstrRegNo)
            If rsTemp.EOF = False Then
                lngBillID = Nvl(rsTemp!ID)
            End If
            
            objHisModule.objModule.zlRefresh mlngPatientId, lngBillID, mlngCurDeptId, Not mblnIsHistory, mblnMoved, 0, True
        
        Case "电子病历"
            If lngPatientFrom = 1 Then
                strSQL = "Select Id From 病人挂号记录 Where No = [1]"
            
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询挂号ID", mstrRegNo)
                If rsTemp.EOF = False Then
                    lngBillID = Nvl(rsTemp!ID)
                End If
            End If
            objHisModule.objModule.zlRefresh mlngPatientId, IIf(lngPatientFrom = 2, mlngPageID, lngBillID), mlngCurDeptId, IIf(lngPatientFrom = 2, 0, 2), IIf(lngPatientFrom = 2, 2, 1)
    End Select
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mobjInAdvice_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
    Call SendMsgToMainWindow(Me, wetViewEprReport, mlngAdviceID, 报告ID & "," & IIf(CanPrint, 1, 0))
End Sub

Private Sub mobjInAdvice_ViewPACSImage(ByVal 医嘱ID As Long)
    Call SendMsgToMainWindow(Me, wetViewPacsImage, 医嘱ID)
End Sub

Private Sub mobjOutAdvice_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
    Call SendMsgToMainWindow(Me, wetViewEprReport, mlngAdviceID, 报告ID & "," & IIf(CanPrint, 1, 0))
End Sub

Private Sub mobjOutAdvice_ViewPACSImage(ByVal 医嘱ID As Long)
    Call SendMsgToMainWindow(Me, wetViewPacsImage, 医嘱ID)
End Sub



