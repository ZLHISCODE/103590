VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkModule_Pathol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IWorkMenu


Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "病理"

Public Enum TModuleType
    mtAll = -1
    mtSpecimen = 0  '标本核收模块
    mtMaterial = 1  '标本取材模块
    mtSlices = 2    '病理制片模块
    mtSpeExam = 3   '病理特检模块
    mtProRep = 4    '病理过程模块
End Enum


Private mlngModule  As Long
Private mstrPrivs  As String
Private mlngCurDeptId As Long
Private mobjOwner As Object

Private mlngAdviceID As Long
Private mlngSendNo As Long
Private mblnMoved As Boolean
Private mlngStudyState As Long

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



'获取菜单接口对象
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


Public Sub NotificationRefresh(Optional ByVal lngModule As Long = -1)
'通知刷新
    Dim i As Long
    Dim objModule As Object
    
    If lngModule = -1 Then
        For i = LBound(aryRegModule) To UBound(aryRegModule)
            If Not aryRegModule(i) Is Nothing Then
                Call aryRegModule(i).NotificationRefresh
            End If
        Next i
    Else
        Set objModule = GetModule(lngModule, False)
        
        If Not objModule Is Nothing Then Call objModule.NotificationRefresh
    End If
    
End Sub

Private Sub Class_Initialize()
    Set mfrmPatholDecalinTask = New frmPatholDecalcification
    Set mfrmConRequest = New frmPatholConsultation
    
    '避免使用ubound时出错，因此默认初始化一个空数组元素
    ReDim Preserve aryRegModule(1)
    Set aryRegModule(0) = Nothing
End Sub

Public Sub Free()
On Error GoTo errHandle
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
    
errHandle:
End Sub


Public Function GetModule(ByVal lngModule As Long, Optional ByVal blnAutoCreate As Boolean = True) As Object
'获取指定模块对象
    Set GetModule = Nothing
    
    Select Case lngModule
        Case mtSpecimen
            If mfrmPatholSpecimen Is Nothing And blnAutoCreate Then
                Set mfrmPatholSpecimen = New frmPatholSpecimen

                Call mfrmPatholSpecimen.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mobjOwner)
                
'                FormSetCaption mfrmPatholSpecimen, False, False
                Call SetWindowStyle(mfrmPatholSpecimen.hWnd)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholSpecimen
            End If
            
            Set GetModule = mfrmPatholSpecimen
            
        Case mtMaterial
            If mfrmPatholMaterial Is Nothing And blnAutoCreate Then
                Set mfrmPatholMaterial = New frmPatholMaterials
                
                Call mfrmPatholMaterial.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mobjOwner)
                
'                FormSetCaption mfrmPatholMaterial, False, False
                Call SetWindowStyle(mfrmPatholMaterial.hWnd)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholMaterial
            End If
            
            Set GetModule = mfrmPatholMaterial
               
        Case mtSlices
            If mfrmPatholSlices Is Nothing And blnAutoCreate Then
                Set mfrmPatholSlices = New frmPatholSlices

                Call mfrmPatholSlices.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mobjOwner)
                
'                FormSetCaption mfrmPatholSlices, False, False
                Call SetWindowStyle(mfrmPatholSlices.hWnd)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholSlices
            End If
            
            Set GetModule = mfrmPatholSlices
            
        Case mtSpeExam
            If mfrmPatholSpeExam Is Nothing And blnAutoCreate Then
                Set mfrmPatholSpeExam = New frmPatholSpecialExamined

                Call mfrmPatholSpeExam.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mobjOwner)
                
'                FormSetCaption mfrmPatholSpeExam, False, False
                Call SetWindowStyle(mfrmPatholSpeExam.hWnd)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholSpeExam
            End If
            
            Set GetModule = mfrmPatholSpeExam
            
        Case mtProRep
            If mfrmPatholProRep Is Nothing And blnAutoCreate Then
                Set mfrmPatholProRep = New frmPatholProcedureRep

                Call mfrmPatholProRep.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mobjOwner)
                
'                FormSetCaption mfrmPatholProRep, False, False
                Call SetWindowStyle(mfrmPatholProRep.hWnd)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholProRep
            End If
            
            Set GetModule = mfrmPatholProRep
    End Select
End Function


Private Sub SetWindowStyle(ByVal lngHandle As Long)
    Dim lngWindowStyle As Long
    
    lngWindowStyle = GetWindowLong(lngHandle, GWL_STYLE)
    
'    If (lngWindowStyle And WS_CHILD) = WS_CHILD Then Exit Sub
    
    lngWindowStyle = lngWindowStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)

    Call SetWindowLong(lngHandle, GWL_STYLE, lngWindowStyle Or WS_CHILD)
End Sub


Public Function GetActiveModule() As Object
'获取当前显示得工作模块
    Dim i As Long
    
    Set GetActiveModule = Nothing
    
    For i = LBound(aryRegModule) To UBound(aryRegModule)
        If Not aryRegModule(i) Is Nothing Then
            If aryRegModule(i).Visible Then
                Set GetActiveModule = aryRegModule(i)
                Exit For
            End If
        End If
    Next i
End Function

'接口实现部分*********************************************************************************

Public Function IWorkMenu_zlIsModuleMenu(ByVal objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单
    Dim blnIsModuleMenu As Boolean
    Dim i As Long
    Dim objIWork As IWorkMenu
    
    blnIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
    
    If Not blnIsModuleMenu Then
        For i = LBound(aryRegModule) To UBound(aryRegModule)
            If Not aryRegModule(i) Is Nothing Then
                Set objIWork = aryRegModule(i)
                
                blnIsModuleMenu = objIWork.zlIsModuleMenu(objControlMenu)
                
                If blnIsModuleMenu Then Exit For
            End If
        Next i
    End If
    
    IWorkMenu_zlIsModuleMenu = blnIsModuleMenu
End Function


Public Sub IWorkMenu_zlCreateMenu(objMenuBar As Object)
'创建影像记录对应的菜单
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim objIWork As IWorkMenu
    
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    
    Set mObjActiveMenuBar = objMenuBar
    
    If Not HasMenu(objMenuBar, conMenu_PatholManage) Then
        Set cbrMenuBar = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholManage, "病理", 3, False)
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
    
    Set objIWork = GetActiveModule
    
    If Not objIWork Is Nothing Then Call objIWork.zlCreateMenu(objMenuBar)
End Sub


Public Sub IWorkMenu_zlCreateToolBar(objToolBar As Object)
'创建工具栏
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim lngIndex As Long
        
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    If HasMenu(objToolBar, conMenu_Pathol_BatProcess) Then Exit Sub
    
    Set cbrLogOut = objToolBar.FindControl(, conMenu_Manage_InQueue)
    
    lngIndex = 4
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index
    
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Pathol_BatProcess, "批量处理", "批量处理制片接受、制片确认、特检接受等操作", G_INT_ICONID_BATPROCESS, True, lngIndex + 1)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Pathol_BatSlicesAccept, "制片批量处理", "制片批量处理", G_INT_ICONID_SLICES, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Pathol_BatSpeExamAccept, "特检批量处理", "特检批量处理", G_INT_ICONID_SPEEXAM, False)
    
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Pathol_Request, "病理申请", "病理申请", 0, False, lngIndex + 2)
    
End Sub

Public Sub IWorkMenu_zlClearMenu()
'清除所创建的菜单
    Dim cbrControl As CommandBarControl
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_PatholManage)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub


Public Sub IWorkMenu_zlClearToolBar()
'清除创建的工具栏
    Dim cbrControl As CommandBarControl
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Pathol_BatProcess)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Pathol_Request)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub

Public Sub IWorkMenu_zlExecuteMenu(ByVal lngMenuId As Long)
'根据菜单ID执行对应功能
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    Dim objIWork As IWorkMenu
    
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
    
    Set objIWork = GetActiveModule
    
    If Not objIWork Is Nothing Then Call objIWork.zlExecuteMenu(lngMenuId)
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim objIWork As IWorkMenu
    
    Select Case control.ID
       Case conMenu_Pathol_Antibody_Manage
            If Not (CheckPopedom(mstrPrivs, "抗体管理") Or CheckPopedom(mstrPrivs, "抗体反馈")) Then
                control.Enabled = False
            Else
            
            End If
        Case conMenu_Pathol_MealManage
            If Not CheckPopedom(mstrPrivs, "套餐维护") Then
                control.Enabled = False
            Else
            
            End If
        Case conMenu_Pathol_Request
            If Not (CheckPopedom(mstrPrivs, "特检申请") Or CheckPopedom(mstrPrivs, "制片申请") Or CheckPopedom(mstrPrivs, "补取申请")) Then
                control.Enabled = False
            Else
                control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_ReportDelay
            If Not CheckPopedom(mstrPrivs, "报告延迟") Then
                control.Enabled = False
            Else
                control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_ConRequest
            If Not CheckPopedom(mstrPrivs, "会诊申请") Then
                control.Enabled = False
            Else
                control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_ConFeedback
            If Not CheckPopedom(mstrPrivs, "会诊反馈") Then
                control.Enabled = False
            Else
                control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_Decalin_Task
            If Not CheckPopedom(mstrPrivs, "病理取材") Then
                control.Enabled = False
            Else
                control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_BatSlicesAccept ', conMenu_Pathol_BatSlicesSure
            If Not CheckPopedom(mstrPrivs, "病理制片") Then
                control.Enabled = False
            Else
'                control.Enabled = Not mblnReadOnly
            End If
            
        Case conMenu_Pathol_BatSpeExamAccept
            If Not (CheckPopedom(mstrPrivs, "免疫组化") Or CheckPopedom(mstrPrivs, "分子病理") Or CheckPopedom(mstrPrivs, "特殊染色")) Then
                control.Enabled = False
            Else
'                control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_Quality_Manage
            If Not CheckPopedom(mstrPrivs, "质量管理") Then
                control.Enabled = False
            Else
                control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_PatholSlices_Quality
            '如果不具备质量管理权限，则可以查看制片质量评审结果
'            If Not CheckPopedom(mstrPrivs, "质量管理") Then
'                control.Enabled = False
'            End If
            control.Enabled = Not mblnReadOnly
        Case conMenu_Pathol_BatProcess
            If Not (CheckPopedom(mstrPrivs, "病理制片") Or CheckPopedom(mstrPrivs, "免疫组化") Or CheckPopedom(mstrPrivs, "分子病理") Or CheckPopedom(mstrPrivs, "特殊染色")) Then
                control.Enabled = False
            Else
'                control.Enabled = Not mblnReadOnly
            End If
    End Select
    
    Set objIWork = GetActiveModule
    
    If Not objIWork Is Nothing Then Call objIWork.zlUpdateMenu(control)
End Sub


Public Sub IWorkMenu_zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单
    Exit Sub
End Sub

Public Sub IWorkMenu_zlRefreshSubMenu(objMenuBar As Object)
'刷新弹出的子菜单
    Exit Sub
End Sub
'*********************************************************************************************


Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing)
'初始化模块参数
    Dim i As Long
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngCurDeptId = lngDepartId
    
    If Not owner Is Nothing Then Set mobjOwner = owner

    For i = LBound(aryRegModule) To UBound(aryRegModule)
        If Not aryRegModule(i) Is Nothing Then
            Call aryRegModule(i).zlInitModule(lngModule, strPrivs, lngDepartId, owner)
        End If
    Next i
End Sub


Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'更新医嘱信息
    
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    mblnReadOnly = False
    
    If Not mfrmPatholSpecimen Is Nothing Then Call mfrmPatholSpecimen.zlUpdateAdviceInf(lngAdviceID, lngSendNO, lngStudyState, blnMoved)
    If Not mfrmPatholMaterial Is Nothing Then Call mfrmPatholMaterial.zlUpdateAdviceInf(lngAdviceID, lngSendNO, lngStudyState, blnMoved)
    If Not mfrmPatholSlices Is Nothing Then Call mfrmPatholSlices.zlUpdateAdviceInf(lngAdviceID, lngSendNO, lngStudyState, blnMoved)
    If Not mfrmPatholSpeExam Is Nothing Then Call mfrmPatholSpeExam.zlUpdateAdviceInf(lngAdviceID, lngSendNO, lngStudyState, blnMoved)
    If Not mfrmPatholProRep Is Nothing Then Call mfrmPatholProRep.zlUpdateAdviceInf(lngAdviceID, lngSendNO, lngStudyState, blnMoved)
    
    '如果为完成，登记，拒绝，审核，则不能对检查进行处理 'Or lngStudyState = 5
    If mlngAdviceID <= 0 Or lngStudyState = 6 Or lngStudyState = 0 Or lngStudyState = 1 Or lngStudyState = -2 Then
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
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Private Sub Menu_Manage_病理号配置()
Dim frmPatholNum As New frmPatholNumConfig
On Error GoTo errFree
    Call frmPatholNum.Show(1, mobjOwner)
    
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
    Call MsgBoxD(mobjOwner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
    Exit Sub
End If


Dim frmAntibody As New frmPatholAntibody
On Error GoTo errFree
    Call frmAntibody.ShowAntibodyManageWind(mstrPrivs, mobjOwner)
errFree:
    Call Unload(frmAntibody)
    Set frmAntibody = Nothing
End Sub



Private Sub Menu_Manage_套餐维护()
'套餐维护

If Not CheckPopedom(mstrPrivs, "套餐维护") Then
    Call MsgBoxD(mobjOwner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
    Exit Sub
End If

Dim frmMeal As New frmPatholMeal
On Error GoTo errFree
    Call frmMeal.ShowMealWindow(mstrPrivs, mobjOwner)
errFree:
    Call Unload(frmMeal)
    Set frmMeal = Nothing
End Sub


Private Sub Menu_Manage_病理申请()
'病理申请
On Error GoTo errFree
    Dim frmRequest As frmPatholRequisition

    If Not (CheckPopedom(mstrPrivs, "特检申请") Or CheckPopedom(mstrPrivs, "制片申请") Or CheckPopedom(mstrPrivs, "补取申请")) Then
        Call MsgBoxD(mobjOwner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If
    
    If mlngAdviceID = 0 Then
        MsgBoxD mobjOwner, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    Set frmRequest = New frmPatholRequisition
    
    Call frmRequest.zlRefresh(mlngAdviceID, mblnReadOnly, mstrPrivs, mblnMoved, mlngCurDeptId, mobjOwner)

    If frmRequest.blnIsUpdate Then
        Call SendMsgToMainWindow(Me, wetPatholRequest, mlngAdviceID)
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
        Call MsgBoxD(mobjOwner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    Set frmDelay = frmPatholReportDelay

    Call frmDelay.zlRefresh(mlngAdviceID, mblnReadOnly, mstrPrivs, mblnMoved, mlngCurDeptId, mobjOwner)
errFree:
    Call Unload(frmDelay)
    Set frmDelay = Nothing
End Sub


Private Sub Menu_Manage_病理质量管理()
'病理质量管理
On Error GoTo errFree
    Dim frmQulity As frmPatholQuality

    If Not CheckPopedom(mstrPrivs, "质量管理") Then
        Call MsgBoxD(mobjOwner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    Set frmQulity = New frmPatholQuality

    Call frmQulity.zlRefresh(mlngAdviceID, mblnReadOnly, mstrPrivs, mblnMoved, mlngCurDeptId, mobjOwner)
    
    If frmQulity.blnIsOk Then
        Call SendMsgToMainWindow(Me, wetPatholQuality, mlngAdviceID, frmQulity.strQuality)
    End If
errFree:
    Call Unload(frmQulity)
    Set frmQulity = Nothing
End Sub

Private Sub Menu_Manage_病理制片质量()
'显示制片质量
On Error GoTo errHandle
     Dim frmSlicesQuality As frmPatholSlices_Quality
    
     If Not CheckPopedom(mstrPrivs, "质量管理") Then
        Call MsgBoxD(mobjOwner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
     End If
     
     Set frmSlicesQuality = New frmPatholSlices_Quality
     
     Call frmSlicesQuality.ShowSlideEvaluateWindow(mlngAdviceID, mlngStudyState, mstrPrivs, mobjOwner)
     
Exit Sub
errHandle:
    Call Unload(frmSlicesQuality)
    Set frmSlicesQuality = Nothing
End Sub


Private Sub Menu_Manage_会诊申请反馈(ByVal lngMenuId As Long)
'会诊申请反馈
On Error GoTo errFree

    If Not (CheckPopedom(mstrPrivs, "会诊申请") Or CheckPopedom(mstrPrivs, "会诊反馈")) Then
        Call MsgBoxD(mobjOwner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    If mlngAdviceID = 0 Then
        MsgBoxD mobjOwner, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If

    Call mfrmConRequest.zlRefresh(mlngAdviceID, mblnReadOnly, mstrPrivs, mblnMoved, mlngCurDeptId, IIf(lngMenuId = conMenu_Pathol_ConFeedback, True, False), mobjOwner)

    Exit Sub
errFree:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_脱钙任务管理()
'脱钙任务管理

If Not CheckPopedom(mstrPrivs, "病理取材") Then
    Call MsgBoxD(mobjOwner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
    Exit Sub
End If

Call mfrmPatholDecalinTask.ShowDecalinTaskWind(mstrPrivs, mblnMoved, mlngModule, mobjOwner)

End Sub



Private Sub Menu_Manage_制片批量接受()
'制片批量接受
On Error GoTo errFree
    Dim frmBatSlices As frmPatholSlices_WorkPrint
  
    If Not CheckPopedom(mstrPrivs, "病理制片") Then
        Call MsgBoxD(mobjOwner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If
    
    Set frmBatSlices = New frmPatholSlices_WorkPrint
    
    Call frmBatSlices.ShowWorkPrint(Nothing, 0, mobjOwner)
        
        
    If frmBatSlices.blnIsOk Then
        Call SendMsgToMainWindow(Me, wetPatholBatSlices, mlngAdviceID)
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
        Call MsgBoxD(mobjOwner, "不具备执行该操作的权限。", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    Set frmBatSpeExam = frmPatholSpecialExamined_WorkPrint
                
    If CheckPopedom(mstrPrivs, "免疫组化") Then
        Call frmBatSpeExam.ShowWorkPrint(Nothing, 0, 0, mstrPrivs, mobjOwner)
    ElseIf CheckPopedom(mstrPrivs, "特殊染色") Then
        Call frmBatSpeExam.ShowWorkPrint(Nothing, 0, 1, mstrPrivs, mobjOwner)
    ElseIf CheckPopedom(mstrPrivs, "分子病理") Then
        Call frmBatSpeExam.ShowWorkPrint(Nothing, 0, 2, mstrPrivs, mobjOwner)
    End If
    
    If frmBatSpeExam.blnIsOk Then
        Call SendMsgToMainWindow(Me, wetPatholBatSpeExm, mlngAdviceID)
    End If
errFree:
    Call Unload(frmBatSpeExam)
    Set frmBatSpeExam = Nothing
End Sub




