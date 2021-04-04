VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkModule_PatholV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IWorkMenuV2


Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "病理"

Public Enum TPatholModuleType
    pmtAll = -1
    pmtSpecimen = 0  '标本核收模块
    pmtMaterial = 1  '标本取材模块
    pmtSlices = 2    '病理制片模块
    pmtSpeExam = 3   '病理特检模块
    pmtProRep = 4    '病理过程模块
End Enum


Private mlngModule  As Long
Private mstrPrivs  As String
Private mlngCurDeptId As Long
Private mObjNotify As IEventNotify

Private mobjStudyInfo As clsStudyInfo
Private mblnReadOnly As Boolean

Private mfrmPatholSpecimen As frmPatholSpecimen              '标本核收
Attribute mfrmPatholSpecimen.VB_VarHelpID = -1
Private mfrmPatholMaterial As frmPatholMaterials             '取材
Attribute mfrmPatholMaterial.VB_VarHelpID = -1
Private mfrmPatholSlices As frmPatholSlices                  '制片
Attribute mfrmPatholSlices.VB_VarHelpID = -1
Private mfrmPatholSpeExam As frmPatholSpecialExamined        '特检
Attribute mfrmPatholSpeExam.VB_VarHelpID = -1
Private mfrmPatholProRep As frmPatholProcedureRep                       '过程报告

Private aryRegModule() As Object                                        '保存当前创建的模块

    
Private mfrmConRequest As frmPatholConsultation                           '会诊申请
Private mfrmPatholDecalinTask As frmPatholDecalcification            '脱钙任务

Private mObjActiveMenuBar As CommandBars

'当前科室ID
Property Get DeptId() As Long
    DeptId = mlngCurDeptId
End Property

'获取菜单接口对象
Property Get zlMenu() As IWorkMenuV2
    Set zlMenu = Me
End Property

'当前医嘱ID
Property Get AdviceId() As Long
    AdviceId = mobjStudyInfo.lngAdviceId
End Property


Private Sub Class_Initialize()
    Set mfrmPatholDecalinTask = New frmPatholDecalcification
    Set mfrmConRequest = New frmPatholConsultation
    
    '避免使用ubound时出错，因此默认初始化一个空数组元素
    ReDim Preserve aryRegModule(0)
    Set aryRegModule(0) = Nothing
End Sub

Public Sub Free()
On Error GoTo errhandle
    Dim i As Long
    
    If SafeArrayGetDim(aryRegModule) <> 0 Then
        For i = LBound(aryRegModule) To UBound(aryRegModule)
            If Not aryRegModule(i) Is Nothing Then
                Unload aryRegModule(i)
                Set aryRegModule(i) = Nothing
            End If
        Next i
    End If
    
    Set mfrmPatholSpecimen = Nothing
    Set mfrmPatholMaterial = Nothing
    Set mfrmPatholSlices = Nothing
    Set mfrmPatholSpeExam = Nothing
    Set mfrmPatholProRep = Nothing
    
    Unload mfrmConRequest
    Set mfrmConRequest = Nothing
    
    Unload mfrmPatholDecalinTask
    Set mfrmPatholDecalinTask = Nothing
    
errhandle:
End Sub


Public Function GetModule(ByVal strModuleName As String, Optional ByVal blnAutoCreate As Boolean = True) As Object
'获取指定模块对象
    Set GetModule = Nothing
    
    Select Case strModuleName
        Case C_TAB_NAME_标本核收
            If mfrmPatholSpecimen Is Nothing And blnAutoCreate Then
                Set mfrmPatholSpecimen = New frmPatholSpecimen

                Call mfrmPatholSpecimen.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mObjNotify.Owner)
                
'                FormSetCaption mfrmPatholSpecimen, False, False
                Call SetWindowStyle(mfrmPatholSpecimen.hwnd, False)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholSpecimen
            End If
            
            Set GetModule = mfrmPatholSpecimen
            
        Case C_TAB_NAME_病理取材
            If mfrmPatholMaterial Is Nothing And blnAutoCreate Then
                Set mfrmPatholMaterial = New frmPatholMaterials
                
                Call mfrmPatholMaterial.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mObjNotify.Owner)
                
                Call SetWindowStyle(mfrmPatholMaterial.hwnd, False)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholMaterial
            End If
            
            Set GetModule = mfrmPatholMaterial
               
        Case C_TAB_NAME_病理制片
            If mfrmPatholSlices Is Nothing And blnAutoCreate Then
                Set mfrmPatholSlices = New frmPatholSlices

                Call mfrmPatholSlices.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mObjNotify.Owner)
                
                Call SetWindowStyle(mfrmPatholSlices.hwnd, False)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholSlices
            End If
            
            Set GetModule = mfrmPatholSlices
            
        Case C_TAB_NAME_病理特检
            If mfrmPatholSpeExam Is Nothing And blnAutoCreate Then
                Set mfrmPatholSpeExam = New frmPatholSpecialExamined

                Call mfrmPatholSpeExam.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mObjNotify.Owner)
                
                Call SetWindowStyle(mfrmPatholSpeExam.hwnd, False)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholSpeExam
            End If
            
            Set GetModule = mfrmPatholSpeExam
            
        Case C_TAB_NAME_过程报告
            If mfrmPatholProRep Is Nothing And blnAutoCreate Then
                Set mfrmPatholProRep = New frmPatholProcedureRep

                Call mfrmPatholProRep.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mObjNotify.Owner)
                
                Call SetWindowStyle(mfrmPatholProRep.hwnd, False)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholProRep
            End If
            
            Set GetModule = mfrmPatholProRep
    End Select
End Function


Private Sub Class_Terminate()
    Set mfrmPatholSpecimen = Nothing
    Set mfrmPatholMaterial = Nothing
    Set mfrmPatholSlices = Nothing
    Set mfrmPatholSpeExam = Nothing
    Set mfrmPatholProRep = Nothing
    
    ReDim aryRegModule(0)
    Set aryRegModule(0) = Nothing
    
    Set mfrmConRequest = Nothing
    Set mfrmPatholDecalinTask = Nothing
    Set mObjActiveMenuBar = Nothing
    Set mObjNotify = Nothing
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
    Dim blnIsModuleMenu As Boolean
    Dim i As Long
    Dim objIWork As IWorkMenuV2
    
    blnIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
    
'    If blnIsModuleMenu = False Then
'        Select Case strModuleName
'            Case "标本核收"
'                If Not mfrmPatholSpecimen Is Nothing Then blnIsModuleMenu = mfrmPatholSpecimen.zlMenu.zlIsModuleMenu(objControlMenu)
'            Case "病理取材"
'                If Not mfrmPatholMaterial Is Nothing Then blnIsModuleMenu = mfrmPatholMaterial.zlMenu.zlIsModuleMenu(objControlMenu)
'            Case "病理制片"
'                If Not mfrmPatholSlices Is Nothing Then blnIsModuleMenu = mfrmPatholSlices.zlMenu.zlIsModuleMenu(objControlMenu)
'            Case "病理特检"
'                If Not mfrmPatholSpeExam Is Nothing Then blnIsModuleMenu = mfrmPatholSpeExam.zlMenu.zlIsModuleMenu(objControlMenu)
'            Case "过程报告"
'                If Not mfrmPatholProRep Is Nothing Then blnIsModuleMenu = mfrmPatholProRep.zlMenu.zlIsModuleMenu(objControlMenu)
'        End Select
'    End If
    
    IWorkMenuV2_zlIsModuleMenu = blnIsModuleMenu
End Function


Public Sub IWorkMenuV2_zlCreateMenu(ByVal strModuleName As String, objMenuBar As Object)
'创建影像记录对应的菜单
    Set mObjActiveMenuBar = objMenuBar
    
    If Len(strModuleName) <= 0 Then
        Call CreatePublicPatholMenu(objMenuBar)
    Else
        Call CreatePrivatePatholMenu(strModuleName, objMenuBar)
    End If
End Sub

Private Sub ClearPrivaePatholMenu(ByVal strModuleName As String, objMenuBar As Object)
    Dim objControlMenu As CommandBarControl
    
    '隐藏之前创建的菜单
    If strModuleName = "标本核收" Then
        Set objControlMenu = objMenuBar.FindControl(, conMenu_PatholSpecimen)
        If Not objControlMenu Is Nothing Then objControlMenu.Delete
    End If
    
    If strModuleName = "病理取材" Then
        Set objControlMenu = objMenuBar.FindControl(, conMenu_PatholMaterial)
        If Not objControlMenu Is Nothing Then objControlMenu.Delete
    End If
    
    If strModuleName = "病理制片" Then
        Set objControlMenu = objMenuBar.FindControl(, conMenu_PatholSlices)
        If Not objControlMenu Is Nothing Then objControlMenu.Delete
    End If
    
    If strModuleName = "病理特检" Then
        Set objControlMenu = objMenuBar.FindControl(, conMenu_PatholSpeExam)
        If Not objControlMenu Is Nothing Then objControlMenu.Delete
    End If
    
    If strModuleName = "过程报告" Then
        Set objControlMenu = objMenuBar.FindControl(, conMenu_PatholProRep)
        If Not objControlMenu Is Nothing Then objControlMenu.Delete
    End If
End Sub

Private Sub CreatePrivatePatholMenu(ByVal strModuleName As String, objMenuBar As Object)
'创建个模块对应的病理菜单
    Call ClearPrivaePatholMenu(strModuleName, objMenuBar)
    
    Select Case strModuleName
        Case "标本核收"
            If Not mfrmPatholSpecimen Is Nothing Then Call mfrmPatholSpecimen.zlMenu.zlCreateMenu(objMenuBar)
        Case "病理取材"
            If Not mfrmPatholMaterial Is Nothing Then Call mfrmPatholMaterial.zlMenu.zlCreateMenu(objMenuBar)
        Case "病理制片"
            If Not mfrmPatholSlices Is Nothing Then Call mfrmPatholSlices.zlMenu.zlCreateMenu(objMenuBar)
        Case "病理特检"
            If Not mfrmPatholSpeExam Is Nothing Then Call mfrmPatholSpeExam.zlMenu.zlCreateMenu(objMenuBar)
        Case "过程报告"
            If Not mfrmPatholProRep Is Nothing Then Call mfrmPatholProRep.zlMenu.zlCreateMenu(objMenuBar)
    End Select
End Sub


Private Sub CreatePublicPatholMenu(objMenuBar As Object)
'创建病理公共菜单
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim objIWork As IWorkMenuV2
    
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    
    If Not HasMenu(objMenuBar, conMenu_PatholManage) Then
        Set cbrMenuBar = objMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholManage, "病理", 3, False)
        cbrMenuBar.ID = conMenu_PatholManage
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_Antibody_Manage, "抗体管理", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_MealManage, "套餐维护", "", 0, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_Decalin_Task, "脱钙任务", "", 0, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_Request, "病理申请", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_ReportDelay, "延迟登记", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_ConRequest, "会诊申请", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_ConFeedback, "会诊反馈", "", 0, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_BatSlicesAccept, "批量制片", "", G_INT_ICONID_SLICES, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_BatSpeExamAccept, "批量特检", "", G_INT_ICONID_SPEEXAM, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_Quality_Manage, "病理质量", "", 0, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSlices_Quality, "玻片质量", "", 3097, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_NumConfig, "号码配置", "", 9023, True)
        End With
    End If
End Sub

Private Sub CreatePublicPatholTools(objToolBar As Object)
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim lngIndex As Long
        
  
    If HasMenu(objToolBar, conMenu_Pathol_BatProcess) Then Exit Sub
    
    Set cbrLogOut = objToolBar.FindControl(, conMenu_Manage_InQueue)
    
    lngIndex = 4
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index
    
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Pathol_BatProcess, "批量处理", "批量处理制片接受、制片确认、特检接受等操作", G_INT_ICONID_BATPROCESS, True, lngIndex + 1)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Pathol_BatSlicesAccept, "制片批量处理", "制片批量处理", G_INT_ICONID_SLICES, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Pathol_BatSpeExamAccept, "特检批量处理", "特检批量处理", G_INT_ICONID_SPEEXAM, False)
    
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Pathol_Request, "病理申请", "病理申请", 0, False, lngIndex + 2)
     
End Sub

Public Sub IWorkMenuV2_zlCreateToolBar(ByVal strModuleName As String, objToolBar As Object)
'创建工具栏
    If Len(strModuleName) <= 0 Then Call CreatePublicPatholTools(objToolBar)
End Sub



Public Sub IWorkMenuV2_zlClearMenu(ByVal strModuleName As String)
'清除所创建的菜单
    Dim cbrControl As CommandBarControl
    
    If mObjActiveMenuBar Is Nothing Then Exit Sub
    
    If Len(strModuleName) <= 0 Then
        Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_PatholManage)
        If Not cbrControl Is Nothing Then Call cbrControl.Delete
    Else
        Call ClearPrivaePatholMenu(strModuleName, mObjActiveMenuBar)
    End If
End Sub


Public Sub IWorkMenuV2_zlClearToolBar(ByVal strModuleName As String)
'清除创建的工具栏
    Dim cbrControl As CommandBarControl
    
    If mObjActiveMenuBar Is Nothing Then Exit Sub
    
    If Len(strModuleName) <= 0 Then
        Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Pathol_BatProcess)
        If Not cbrControl Is Nothing Then Call cbrControl.Delete
        
        Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Pathol_Request)
        If Not cbrControl Is Nothing Then Call cbrControl.Delete
    End If
End Sub

Public Sub IWorkMenuV2_zlExecuteMenu(ByVal strModuleName As String, ByVal lngMenuId As Long)
'根据菜单ID执行对应功能
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    Dim objIWork As IWorkMenuV2
    
    Select Case lngMenuId
        Case conMenu_Pathol_NumConfig     '病理号码配置
            Call Menu_Manage_病理号配置
            
        Case conMenu_Pathol_Antibody_Manage    '抗体管理
            Call Menu_Manage_抗体管理
            
        Case conMenu_Pathol_MealManage        '套餐维护
            Call Menu_Manage_套餐维护
            
        Case conMenu_Pathol_Request     '病理申请
            Call Menu_Manage_病理申请
            
        Case conMenu_Pathol_ReportDelay       '延迟登记
            Call Menu_Manage_延迟登记
        
        Case conMenu_Pathol_ConRequest, conMenu_Pathol_ConFeedback       '会诊申请反馈
            Call Menu_Manage_会诊申请反馈(lngMenuId)
            
        Case conMenu_Pathol_Decalin_Task       '脱钙任务
            Call Menu_Manage_脱钙任务管理
            
        Case conMenu_Pathol_BatSlicesAccept   '制片批量接受
            Call Menu_Manage_制片批量接受
            
        Case conMenu_Pathol_BatSpeExamAccept  '特检批量接受
            Call Menu_Manage_特检批量接受
            
        Case conMenu_Pathol_Quality_Manage '病理质量管理
            Call Menu_Manage_病理质量管理
            
         Case conMenu_PatholSlices_Quality  '病理制片质量
            Call Menu_Manage_病理制片质量
    End Select
End Sub


Public Sub IWorkMenuV2_zlUpdateMenu(ByVal strModuleName As String, Control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim objIWork As IWorkMenuV2
    
    Select Case Control.ID
       Case conMenu_PatholManage
            Control.Enabled = True
        
       Case conMenu_Pathol_Antibody_Manage
            If Not (CheckPopedom(mstrPrivs, "抗体管理") Or CheckPopedom(mstrPrivs, "抗体反馈")) Then
                Control.Enabled = False
            Else
            
            End If
        Case conMenu_Pathol_MealManage
            If Not CheckPopedom(mstrPrivs, "套餐维护") Then
                Control.Enabled = False
            Else
            
            End If
        Case conMenu_Pathol_Request
            If Not (CheckPopedom(mstrPrivs, "特检申请") Or CheckPopedom(mstrPrivs, "制片申请") Or CheckPopedom(mstrPrivs, "补取申请")) Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_ReportDelay
            If Not CheckPopedom(mstrPrivs, "报告延迟") Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_ConRequest
            If Not CheckPopedom(mstrPrivs, "会诊申请") Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_ConFeedback
            If Not CheckPopedom(mstrPrivs, "会诊反馈") Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_Decalin_Task
            If Not CheckPopedom(mstrPrivs, "病理取材") Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_BatSlicesAccept ', conMenu_Pathol_BatSlicesSure
            If Not CheckPopedom(mstrPrivs, "病理制片") Then
                Control.Enabled = False
            Else
'                control.Enabled = Not mblnReadOnly
            End If
            
        Case conMenu_Pathol_BatSpeExamAccept
            If Not (CheckPopedom(mstrPrivs, "免疫组化") Or CheckPopedom(mstrPrivs, "分子病理") Or CheckPopedom(mstrPrivs, "特殊染色")) Then
                Control.Enabled = False
            Else
'                control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_Quality_Manage
            If Not CheckPopedom(mstrPrivs, "质量管理") Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_PatholSlices_Quality
            '如果不具备质量管理权限，则可以查看制片质量评审结果
'            If Not CheckPopedom(mstrPrivs, "质量管理") Then
'                control.Enabled = False
'            End If
            Control.Enabled = Not mblnReadOnly
        Case conMenu_Pathol_BatProcess
            If Not (CheckPopedom(mstrPrivs, "病理制片") Or CheckPopedom(mstrPrivs, "免疫组化") Or CheckPopedom(mstrPrivs, "分子病理") Or CheckPopedom(mstrPrivs, "特殊染色")) Then
                Control.Enabled = False
            Else
'                control.Enabled = Not mblnReadOnly
            End If
    End Select
End Sub


Public Sub IWorkMenuV2_zlPopupMenu(ByVal strModuleName As String, objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单
    Exit Sub
End Sub

Public Sub IWorkMenuV2_zlRefreshSubMenu(ByVal strModuleName As String, objMenuBar As Object)
'刷新弹出的子菜单
    Exit Sub
End Sub
'*********************************************************************************************


Public Sub zlInitModule(objNotify As IEventNotify, ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long)
'初始化模块参数
    Dim i As Long
    
    Set mObjNotify = objNotify
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngCurDeptId = lngDepartId
     

    For i = LBound(aryRegModule) To UBound(aryRegModule)
        If Not aryRegModule(i) Is Nothing Then
            Call aryRegModule(i).zlInitModule(lngModule, strPrivs, lngDepartId, mObjNotify.Owner)
        End If
    Next i
End Sub


Public Sub zlRefresh(objStudyInfo As clsStudyInfo, ByVal strSelModuleName As String, _
    Optional ByVal blnIsForceRefresh As Boolean = False, Optional ByVal blnIsHistory As Boolean = False)
'更新医嘱信息
    
    If Not mobjStudyInfo Is Nothing And Not objStudyInfo Is Nothing Then
        If mobjStudyInfo.IsEquals(objStudyInfo) And blnIsForceRefresh = False Then Exit Sub
    End If
    
    Set mobjStudyInfo = objStudyInfo

    mblnReadOnly = False
    
    Select Case strSelModuleName
        Case C_TAB_NAME_标本核收
            Call mfrmPatholSpecimen.zlUpdateAdviceInf(mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.intStep, mobjStudyInfo.blnMoved)
            Call mfrmPatholSpecimen.zlRefreshFace(blnIsForceRefresh)
            
        Case C_TAB_NAME_病理取材
            Call mfrmPatholMaterial.zlUpdateAdviceInf(mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.intStep, mobjStudyInfo.blnMoved)
            Call mfrmPatholMaterial.zlRefreshFace(blnIsForceRefresh)
            
        Case C_TAB_NAME_病理制片
            Call mfrmPatholSlices.zlUpdateAdviceInf(mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.intStep, mobjStudyInfo.blnMoved)
            Call mfrmPatholSlices.zlRefreshFace(blnIsForceRefresh)
            
        Case C_TAB_NAME_病理特检
            Call mfrmPatholSpeExam.zlUpdateAdviceInf(mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.intStep, mobjStudyInfo.blnMoved)
            Call mfrmPatholSpeExam.zlRefreshFace(blnIsForceRefresh)
            
        Case C_TAB_NAME_过程报告
            Call mfrmPatholProRep.zlUpdateAdviceInf(mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.intStep, mobjStudyInfo.blnMoved)
            Call mfrmPatholProRep.zlRefreshFace(blnIsForceRefresh)
            
    End Select
    
    '如果为完成，登记，拒绝，审核，则不能对检查进行处理 'Or lngStudyState = 5
    If mobjStudyInfo.lngAdviceId <= 0 Or mobjStudyInfo.intStep = 6 Or mobjStudyInfo.intStep = 0 Or mobjStudyInfo.intStep = 1 Or mobjStudyInfo.intStep = -2 Then
        mblnReadOnly = True
    End If
End Sub



Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, _
    Optional blnStartGroup As Boolean = False, Optional lngIndex As Long) As CommandBarControl
'创建该模块内的菜单
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '如果这里不指定id，则不能将有些菜单添加到右键菜单中
    
    If lngIconId <> 0 Then CreateModuleMenu.iconid = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Private Sub Menu_Manage_病理号配置()
Dim frmPatholNum As New frmPatholNumConfig
On Error GoTo errFree
    Call frmPatholNum.Show(1, mObjNotify.Owner)
    
errFree:
    Call Unload(frmPatholNum)
    Set frmPatholNum = Nothing
End Sub

Public Function GetWindowCaption() As String
    GetWindowCaption = "病理检查系统"
End Function

Private Sub Menu_Manage_抗体管理()
'抗体管理
If Not (CheckPopedom(mstrPrivs, "抗体管理") Or CheckPopedom(mstrPrivs, "抗体反馈")) Then
    Call MsgBoxD(mObjNotify.Owner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
    Exit Sub
End If


Dim frmAntibody As New frmPatholAntibody
On Error GoTo errFree
    Call frmAntibody.ShowAntibodyManageWind(mstrPrivs, mObjNotify.Owner)
errFree:
    Call Unload(frmAntibody)
    Set frmAntibody = Nothing
End Sub



Private Sub Menu_Manage_套餐维护()
'套餐维护

If Not CheckPopedom(mstrPrivs, "套餐维护") Then
    Call MsgBoxD(mObjNotify.Owner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
    Exit Sub
End If

Dim frmMeal As New frmPatholMeal
On Error GoTo errFree
    Call frmMeal.ShowMealWindow(mstrPrivs, mObjNotify.Owner)
errFree:
    Call Unload(frmMeal)
    Set frmMeal = Nothing
End Sub


Private Sub Menu_Manage_病理申请()
'病理申请
On Error GoTo errFree
    Dim frmRequest As frmPatholRequisition

    If Not (CheckPopedom(mstrPrivs, "特检申请") Or CheckPopedom(mstrPrivs, "制片申请") Or CheckPopedom(mstrPrivs, "补取申请")) Then
        Call MsgBoxD(mObjNotify.Owner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD mObjNotify.Owner, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    Set frmRequest = New frmPatholRequisition
    
    Call frmRequest.zlRefresh(mobjStudyInfo.lngAdviceId, mblnReadOnly, mstrPrivs, mobjStudyInfo.blnMoved, mlngCurDeptId, mObjNotify.Owner)

    If frmRequest.blnIsUpdate Then
        Call SendMsgToMainWindow(Me, wetPatholRequest, mobjStudyInfo.lngAdviceId)
    End If
errFree:
    Call Unload(frmRequest)
    Set frmRequest = Nothing
End Sub


Private Sub Menu_Manage_延迟登记()
'延迟登记
On Error GoTo errFree
    Dim frmDelay As frmPatholReportDelay

    If Not CheckPopedom(mstrPrivs, "报告延迟") Then
        Call MsgBoxD(mObjNotify.Owner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    Set frmDelay = frmPatholReportDelay

    Call frmDelay.zlRefresh(mobjStudyInfo.lngAdviceId, mblnReadOnly, mstrPrivs, mobjStudyInfo.blnMoved, mlngCurDeptId, mObjNotify.Owner)
errFree:
    Call Unload(frmDelay)
    Set frmDelay = Nothing
End Sub


Private Sub Menu_Manage_病理质量管理()
'病理质量管理
On Error GoTo errFree
    Dim frmQulity As frmPatholQuality

    If Not CheckPopedom(mstrPrivs, "质量管理") Then
        Call MsgBoxD(mObjNotify.Owner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    Set frmQulity = New frmPatholQuality

    Call frmQulity.zlRefresh(mobjStudyInfo.lngAdviceId, mblnReadOnly, mstrPrivs, mobjStudyInfo.blnMoved, mlngCurDeptId, mObjNotify.Owner)
    
    If frmQulity.blnIsOk Then
        Call SendMsgToMainWindow(Me, wetPatholQuality, mobjStudyInfo.lngAdviceId, frmQulity.strQuality)
    End If
errFree:
    Call Unload(frmQulity)
    Set frmQulity = Nothing
End Sub

Private Sub Menu_Manage_病理制片质量()
'显示制片质量
On Error GoTo errhandle
     Dim frmSlicesQuality As frmPatholSlices_Quality
    
     If Not CheckPopedom(mstrPrivs, "质量管理") Then
        Call MsgBoxD(mObjNotify.Owner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
     End If
     
     Set frmSlicesQuality = New frmPatholSlices_Quality
     
     Call frmSlicesQuality.ShowSlideEvaluateWindow(mobjStudyInfo.lngAdviceId, mobjStudyInfo.intStep, mstrPrivs, mObjNotify.Owner)
     
Exit Sub
errhandle:
    Call Unload(frmSlicesQuality)
    Set frmSlicesQuality = Nothing
End Sub


Private Sub Menu_Manage_会诊申请反馈(ByVal lngMenuId As Long)
'会诊申请反馈
On Error GoTo errFree

    If Not (CheckPopedom(mstrPrivs, "会诊申请") Or CheckPopedom(mstrPrivs, "会诊反馈")) Then
        Call MsgBoxD(mObjNotify.Owner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD mObjNotify.Owner, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If

    Call mfrmConRequest.zlRefresh(mobjStudyInfo.lngAdviceId, mblnReadOnly, mstrPrivs, mobjStudyInfo.blnMoved, mlngCurDeptId, IIf(lngMenuId = conMenu_Pathol_ConFeedback, True, False), mObjNotify.Owner)

    Exit Sub
errFree:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_脱钙任务管理()
'脱钙任务管理

If Not CheckPopedom(mstrPrivs, "病理取材") Then
    Call MsgBoxD(mObjNotify.Owner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
    Exit Sub
End If

Call mfrmPatholDecalinTask.ShowDecalinTaskWind(mstrPrivs, mobjStudyInfo.blnMoved, mlngModule, mObjNotify.Owner)

End Sub



Private Sub Menu_Manage_制片批量接受()
'制片批量接受
On Error GoTo errFree
    Dim frmBatSlices As frmPatholSlices_WorkPrint
  
    If Not CheckPopedom(mstrPrivs, "病理制片") Then
        Call MsgBoxD(mObjNotify.Owner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If
    
    Set frmBatSlices = New frmPatholSlices_WorkPrint
    
    Call frmBatSlices.ShowWorkPrint(Nothing, 0, mObjNotify.Owner)
        
        
    If frmBatSlices.blnIsOk Then
        Call SendMsgToMainWindow(Me, wetPatholBatSlices, mobjStudyInfo.lngAdviceId)
    End If
errFree:
    Call Unload(frmBatSlices)
    Set frmBatSlices = Nothing
End Sub



Private Sub Menu_Manage_特检批量接受()
'特检批量接受
On Error GoTo errFree
    Dim frmBatSpeExam As frmPatholSpecialExamined_WorkPrint
    
    If Not (CheckPopedom(mstrPrivs, "免疫组化") Or CheckPopedom(mstrPrivs, "特殊染色") Or CheckPopedom(mstrPrivs, "分子病理")) Then
        Call MsgBoxD(mObjNotify.Owner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    Set frmBatSpeExam = frmPatholSpecialExamined_WorkPrint
                
    If CheckPopedom(mstrPrivs, "免疫组化") Then
        Call frmBatSpeExam.ShowWorkPrint(Nothing, 0, 0, mstrPrivs, mObjNotify.Owner)
    ElseIf CheckPopedom(mstrPrivs, "特殊染色") Then
        Call frmBatSpeExam.ShowWorkPrint(Nothing, 0, 1, mstrPrivs, mObjNotify.Owner)
    ElseIf CheckPopedom(mstrPrivs, "分子病理") Then
        Call frmBatSpeExam.ShowWorkPrint(Nothing, 0, 2, mstrPrivs, mObjNotify.Owner)
    End If
    
    If frmBatSpeExam.blnIsOk Then
        Call SendMsgToMainWindow(Me, wetPatholBatSpeExm, mobjStudyInfo.lngAdviceId)
    End If
errFree:
    Call Unload(frmBatSpeExam)
    Set frmBatSpeExam = Nothing
End Sub
 
