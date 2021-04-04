VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkModule_ReportV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IWorkMenuV2


Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "报告"
Private Const M_LNG_MENU_BASE_ID As Long = 1000000

'Private mReportType As ReportType '报告编辑器类型

Private mlngModule  As Long
Private mstrPrivs  As String
Private mlngDeptID As Long
Private mObjNotify As IEventNotify

Private mobjCapLinker As ICapHelper
Private mobjMainHelper As Object

Private mobjStudyInfo As clsStudyInfo
Private mlngBabyNum As Long
 
Private mObjActiveMenuBar As CommandBars

Private WithEvents mobjEmrReport As zlRichEPR.cDockReport  '电子病历报告对象
Attribute mobjEmrReport.VB_VarHelpID = -1
Private mobjPacsReport As frmReportV2     'PACS报告编辑器，嵌入主程序的窗口
Attribute mobjPacsReport.VB_VarHelpID = -1
Private mobjRichReportView As frmRichReportV2
 
 
Private mobjStudyList As Object
Private mblnIsShowing As Boolean
Private mblnIsNoAskPrint As Boolean

'报告相关参数
Private mblnTechReptSame As Boolean
Private mblnIsEditWithRepImage As Boolean

Private mstrDoDoctor As String  '检查技师
Private mblnHasReport As Boolean

Private mstrReportLevel As String    '报告质量等级串
Private mintReportLevel As Integer   '报告质量判定
Private mblnIsProcessing As Boolean

Private mblnIsRefreshStudy As Boolean


'科室ID
Property Get DeptId() As Long
    DeptId = mlngDeptID
End Property

'检查信息
Property Get StudyInfo() As clsStudyInfo
    Set StudyInfo = mobjStudyInfo
End Property

Property Set StudyInfo(value As clsStudyInfo)
    Set mobjStudyInfo = value
    
    
    
    mblnIsRefreshStudy = False
End Property

Property Get AdviceId() As Long
    AdviceId = mobjStudyInfo.lngAdviceId
End Property

Property Get AlreadyAdviceId() As Long
    If mblnIsRefreshStudy Then
        AlreadyAdviceId = mobjStudyInfo.lngAdviceId
    Else
        AlreadyAdviceId = 0
    End If
End Property


Property Get BaseMenuId() As Long
    BaseMenuId = M_LNG_MENU_BASE_ID
End Property

Property Get Processing() As Boolean
    Processing = mblnIsProcessing
End Property

'获取菜单接口对象
Property Get zlMenu() As IWorkMenuV2
    Set zlMenu = Me
End Property

''编辑器类型
'Property Get EditorType() As Long
'    EditorType = mReportType
'End Property


'获取报告窗口句柄
Property Get zlGetForm(ByVal strModuleName As String) As Object
    Select Case strModuleName
        Case C_WORKMODULE_NAME_老版报告
            If mobjPacsReport Is Nothing Then
                Set mobjPacsReport = New frmReportV2
                
                Call SetWindowStyle(mobjPacsReport.hwnd, False)
                
                Call mobjPacsReport.zlInit(mObjNotify, mlngModule, mlngDeptID, mstrPrivs, mobjCapLinker, mobjMainHelper)
            End If
            
            Set zlGetForm = mobjPacsReport
        Case C_WORKMODULE_NAME_病历报告
            If mobjEmrReport Is Nothing Then
                Set mobjEmrReport = New zlRichEPR.cDockReport
            End If
            
            Set zlGetForm = mobjEmrReport.zlGetForm()
        Case Else
            'pacs智能文档编辑器，仅供查看
            If mobjRichReportView Is Nothing Then
                Set mobjRichReportView = New frmRichReportV2
                
                Call SetWindowStyle(mobjRichReportView.hwnd, False)
                
                Call mobjRichReportView.zlInit(mObjNotify, mlngModule, mlngDeptID, mstrPrivs)
            End If
            
            Set zlGetForm = mobjRichReportView
    End Select
End Property


Public Sub ReInit(ByVal strModuleName As String)
    Select Case strModuleName
        Case C_WORKMODULE_NAME_老版报告
            If Not mobjPacsReport Is Nothing Then Call mobjPacsReport.zlInit(mObjNotify, mlngModule, mlngDeptID, mstrPrivs, mobjCapLinker, mobjMainHelper)
  
        Case C_WORKMODULE_NAME_病历报告

        Case Else
            'pacs智能文档编辑器，仅供查看
            If Not mobjRichReportView Is Nothing Then Call mobjRichReportView.zlInit(mObjNotify, mlngModule, mlngDeptID, mstrPrivs)
            
    End Select
End Sub

Private Sub Class_Initialize()
'    mReportType = PACS报告编辑器               '报告编辑器

    mblnIsNoAskPrint = False
End Sub

Public Sub Free()
On Error GoTo errhandle
    
    Call DisposeCurReportObj
errhandle:
End Sub



Private Sub Class_Terminate()
    If Not mobjPacsReport Is Nothing Then Unload mobjPacsReport
    If Not mobjEmrReport Is Nothing Then Unload mobjEmrReport.zlGetForm
    
    Set mObjNotify = Nothing
    Set mObjActiveMenuBar = Nothing
    
    Set mobjStudyList = Nothing
    
    Set mobjEmrReport = Nothing
    Set mobjPacsReport = Nothing
    
    Set mobjMainHelper = Nothing
    
    Set mobjStudyInfo = Nothing
End Sub

'接口实现部分*********************************************************************************

Public Function IWorkMenuV2_zlBaseMenuID() As Long
    IWorkMenuV2_zlBaseMenuID = M_LNG_MENU_BASE_ID
End Function

Public Function IWorkMenuV2_zlExecuteCmd(ByVal lngCmdType As Long)
'执行菜单命令

End Function


Public Function IWorkMenuV2_zlIsModuleMenu(ByVal strModuleName As String, objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单
'医嘱，费用，病历模块创建的菜单名对应为“医嘱”，“费用”，“病历”
    
    IWorkMenuV2_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenuV2_zlCreateMenu(ByVal strModuleName As String, objMenuBar As Object)
'创建影像记录对应的菜单
    Dim objNewMenu As CommandBarControl
    Dim objWrite As CommandBarControl
    Dim objAudit As CommandBarControl
    
    Set mObjActiveMenuBar = objMenuBar
    
    Set objNewMenu = objMenuBar.FindControl(, conMenu_EditPopup)

    '如果菜单存在，则先删除对应的菜单和工具栏按钮重新创建
    Call DelModuleMenuAndToolBar(objMenuBar, objNewMenu, M_STR_MODULE_MENU_TAG)

    Select Case strModuleName
        Case C_WORKMODULE_NAME_老版报告
            Call mobjPacsReport.zlMenu.zlCreateMenu("", objMenuBar)
        Case C_WORKMODULE_NAME_病历报告
            Call mobjEmrReport.zlDefCommandBars(mObjActiveMenuBar)
        
            '需要隐藏工具栏的书写按钮，因为主界面创建了一个通用报告书写按钮
            Set objWrite = mObjActiveMenuBar(2).Controls.Find(, conMenu_Edit_Modify, , False)
            If Not objWrite Is Nothing Then
                objWrite.Visible = False
                
    '            Set objAudit = mObjActiveMenuBar(2).Controls.Find(, conMenu_Edit_Audit, , False)
    '            If Not objAudit Is Nothing Then   objAudit..Index = objWrite.Index + 1
            End If
            
            Set objAudit = mObjActiveMenuBar(2).Controls.Find(, conMenu_Edit_Audit, , False)
            If Not objAudit Is Nothing Then
                objAudit.Visible = False
            End If
        Case Else
            'Pacs智能文档编辑器，不创建相关项
    End Select
    
    
    '添加预览、打印、批量打印的菜单栏和工具栏按钮
    Call zlAttachPublicMenu(mObjActiveMenuBar)
        
    '创建报告质量菜单
    Call CreateQualityMenu
    
    
    Call SetMenuCategory(objMenuBar, objNewMenu, M_STR_MODULE_MENU_TAG)
    
    '避免创建模块工具栏时，窗口左上角出现不完全显示的图标，一闪而过的现象
'    Call objMenuBar.RecalcLayout
End Sub

Private Sub zlAttachPublicMenu(objMenuBar As Object)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    
    
    '创建报告打印预览等基本菜单
    Set cbrMenuBar = objMenuBar.ActiveMenuBar.Controls.Find(, conMenu_EditPopup)
    If cbrMenuBar Is Nothing Then Exit Sub
    
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_NoAskPrint, "使用静默打印", "", 0, False, 1)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Preview, "预览", "", 102, False, 2)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Print, "打印", "", 103, False, 3)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_BatPrint, "批量打印", "", 122, False, 4)
        If cbrMenuBar.CommandBar.Controls.Count >= 5 Then cbrMenuBar.CommandBar.Controls(5).BeginGroup = True
    End With
    
End Sub

Private Sub CreateSubordinateMenuTools(ByVal strReportLevel As String, ByVal cbrControl As CommandBarControl)
'创建下级菜单和工具栏
    Dim cbrPopControl As CommandBarControl
    Dim intTxtLen As Integer
    Dim i As Integer
    
    intTxtLen = Len(strReportLevel) - Len(Replace(strReportLevel, ",", "")) + 1
    For i = 1 To 4
        If i <= intTxtLen Then
            If Trim(Split(strReportLevel, ",")(i - 1)) <> "" Then
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, _
                    Decode(i, 1, conMenu_Manage_ReportFirst, 2, conMenu_Manage_ReportSecond, 3, conMenu_Manage_ReportThird, 4, conMenu_Manage_ReportFourth), Trim(Split(strReportLevel, ",")(i - 1)), "", 0, False)
                cbrPopControl.Category = "" ' M_STR_MODULE_MENU_TAG
            End If
        End If
    Next i
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
    
    CreateModuleMenu.Category = "" 'M_STR_MODULE_MENU_TAG
End Function

Private Sub SetMenuCategory(objMenuBar As Object, objMenu As CommandBarControl, ByVal strCategory As String)
'配置Category属性
    Dim objToolBar As CommandBar
    Dim objActiveMenu As CommandBar
    Dim objSubMenu As CommandBar
    Dim lngIconId As Long
    
    Dim i As Long
    Dim j As Long

    Set objActiveMenu = objMenuBar.ActiveMenuBar
    For i = 1 To objActiveMenu.Controls.Count
        If objActiveMenu.Controls(i).type <> xtpControlButton Then
            Set objSubMenu = objActiveMenu.Controls(i).CommandBar
            
            For j = objSubMenu.Controls.Count To 1 Step -1
                If objSubMenu.Controls(j).Category = "" Then
                    objSubMenu.Controls(j).Category = strCategory
                    
                    lngIconId = objSubMenu.Controls(j).iconid
                    
                    objSubMenu.Controls(j).ID = objSubMenu.Controls(j).ID + M_LNG_MENU_BASE_ID
                    objSubMenu.Controls(j).iconid = lngIconId
                End If
            Next j
            
            If objActiveMenu.Controls(i).Category = "" Then
                objActiveMenu.Controls(i).Category = strCategory
                
                lngIconId = objActiveMenu.Controls(i).iconid
                
                objActiveMenu.Controls(i).ID = objActiveMenu.Controls(i).ID + M_LNG_MENU_BASE_ID
                objActiveMenu.Controls(i).iconid = lngIconId
            End If
        End If
    Next i
    
    Set objToolBar = objMenuBar.Item(2)
    If Not objToolBar Is Nothing Then
        For i = 1 To objToolBar.Controls.Count
            If objToolBar.Controls(i).Category = "" Then
                objToolBar.Controls(i).Category = strCategory
                
                '避免因修改id造成图标不能显示
                lngIconId = objToolBar.Controls(i).iconid
                
                objToolBar.Controls(i).ID = objToolBar.Controls(i).ID + M_LNG_MENU_BASE_ID
                objToolBar.Controls(i).iconid = lngIconId
            End If
            
            If objToolBar.Controls(i).type <> xtpControlButton And objToolBar.Controls(i).type <> xtpControlButtonPopup Then
                Set objSubMenu = objToolBar.Controls(i).CommandBar
                
                For j = objSubMenu.Controls.Count To 1 Step -1
                    If objSubMenu.Controls(j).Category = "" Then
                        objSubMenu.Controls(j).Category = strCategory
                        
                        lngIconId = objSubMenu.Controls(j).iconid
                        
                        objSubMenu.Controls(j).ID = objSubMenu.Controls(j).ID + M_LNG_MENU_BASE_ID
                        objSubMenu.Controls(j).iconid = lngIconId
                    End If
                Next j
            End If
        Next i
    End If
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


Public Sub IWorkMenuV2_zlCreateToolBar(ByVal strModuleName As String, objToolBar As Object)
'创建工具栏
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    
    Select Case strModuleName
        Case C_WORKMODULE_NAME_老版报告
    '        Call mobjPacsReport.zlMenu.zlCreateToolBar(objToolBar)
        
        Case C_WORKMODULE_NAME_病历报告
    '        Set cbrToolBar = mObjActiveMenuBar(2)
    '
    '        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Preview, "预览", "", 102, True, 0)
    '        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Print, "打印", "", 103, False, 1)
        Case Else
             'PACS智能文档编辑器不创建相关项
        
    End Select
    
    '创建报告质量按钮
    Call CreateQualityToolBar(objToolBar)
    
    Call SetMenuCategory(mObjActiveMenuBar, Nothing, M_STR_MODULE_MENU_TAG)

End Sub

Private Sub CreateQualityMenu()
'创建质量菜单
    Dim objNewMenu As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim objGC As CommandBarControl
    
    '创建报告质量菜单
    Set cbrMenuBar = mObjActiveMenuBar.FindControl(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then Exit Sub
    
    With cbrMenuBar.CommandBar
        Set objGC = cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_ReportQuality, True, True)
        If Not objGC Is Nothing Then
            objGC.Delete
        End If

        If CheckPopedom(mstrPrivs, "报告质控") Then
            Set objGC = cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_GChannel, , True)
            If objGC Is Nothing Then
                Set objNewMenu = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_ReportQuality, "报告质量", "", 8346, False, .Controls.Count - 1)
            Else
                Set objNewMenu = CreateModuleMenu(objGC.Parent.Controls, xtpControlPopup, conMenu_Manage_ReportQuality, "报告质量", "", 8346, False, objGC.Index - 1)
            End If

            objNewMenu.Category = ""
            Call CreateSubordinateMenuTools(mstrReportLevel, objNewMenu)
        End If

    End With
End Sub

Private Sub CreateQualityToolBar(objToolBar As Object)
'创建质量工具栏
    Dim cbrControl As CommandBarControl
    
    '创建报告质量工具栏
    If CheckPopedom(mstrPrivs, "报告质控") Then
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_ReportQuality, "报告质量", "报告质量", 8346, False, IIf(mlngModule <> G_LNG_PATHOLSYS_NUM, mObjActiveMenuBar.FindControl(, conMenu_Manage_Complete).Index - 1, mObjActiveMenuBar.FindControl(, conMenu_Manage_Complete).Index))
        cbrControl.Category = "" 'M_STR_MODULE_MENU_TAG
        Call CreateSubordinateMenuTools(mstrReportLevel, cbrControl)
    End If
End Sub

    
Public Sub IWorkMenuV2_zlClearMenu(ByVal strModuleName As String)
'清除所创建的菜单
    Dim cbrControl As CommandBarControl

    If mObjActiveMenuBar Is Nothing Then Exit Sub

    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_EditPopup + M_LNG_MENU_BASE_ID)
'    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    Call DelModuleMenuAndToolBar(mObjActiveMenuBar, cbrControl, M_STR_MODULE_MENU_TAG)
End Sub


Public Sub IWorkMenuV2_zlClearToolBar(ByVal strModuleName As String)
'清除创建的工具栏

    '菜单删除时，已同步删除了对应的工具栏
    If mObjActiveMenuBar Is Nothing Then Exit Sub
    
    Exit Sub
End Sub



Public Sub IWorkMenuV2_zlExecuteMenu(ByVal strModuleName As String, ByVal lngMenuId As Long)
'根据菜单ID执行对应功能
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    Dim lngCurMenuId As Long
    Dim lngIconId As Long
    Dim blnIsCancel As Boolean
    Dim lngQuality As Long
    Dim blnNewMenu As Boolean
    
    lngCurMenuId = lngMenuId - M_LNG_MENU_BASE_ID
    
    '批量打印
    Select Case lngCurMenuId
        Case conMenu_File_BatPrint
            
            Call BatPrint
        
        Case conMenu_File_NoAskPrint '静默打印
            Set objCbrControl = mObjActiveMenuBar.FindControl(, lngCurMenuId + M_LNG_MENU_BASE_ID, , True)
            If Not objCbrControl Is Nothing Then
                objCbrControl.Checked = Not objCbrControl.Checked
                zlDatabase.SetPara "NoAsk", IIf(objCbrControl.Checked, 1, 0), glngSys, 1070
                
                mblnIsNoAskPrint = objCbrControl.Checked
            End If
    
        Case conMenu_Manage_ReportFirst, conMenu_Manage_ReportSecond, conMenu_Manage_ReportThird, conMenu_Manage_ReportFourth '四级报告质量
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_QUALITY, 0, mobjStudyInfo.lngAdviceId, 0, blnIsCancel)
            If blnIsCancel Then Exit Sub
            
            lngQuality = SetReportQuality(lngCurMenuId, mstrReportLevel)
            If lngQuality <> 0 Then
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_QUALITY, 1, mobjStudyInfo.lngAdviceId, lngQuality)
            End If
            
            
        Case Else
        
           '控制 只能书写自己检查的报告（检查技师需要和当前登录用户相同）,'不允许书写、修订、删除
            If mblnTechReptSame = True And mobjStudyInfo.strDoDoctor <> "" And mobjStudyInfo.strDoDoctor <> UserInfo.姓名 Then
                Call MsgBoxD(mObjNotify.Owner, "你不是这个患者的检查技师，无法操作这份报告。", vbInformation, "提示")
                Exit Sub
            End If
            
            If mblnIsEditWithRepImage = True And Len(mobjStudyInfo.strStudyUID) <= 0 And lngCurMenuId <> conMenu_Edit_Delete Then
'                If MsgBoxD(mObjNotify.Owner, "本次检查未找到相关图像，是否强制书写。", vbYesNo, "提示") = vbNo Then Exit Sub
                Call MsgBoxD(mObjNotify.Owner, "本次检查未找到相关图像，不能书写。", vbOKOnly, "提示")
                Exit Sub
            End If
            
            
            If lngCurMenuId = conMenu_Edit_Audit _
                Or lngCurMenuId = conMenu_Edit_Modify _
                Or lngCurMenuId = conMenu_PacsReport_Open _
                Or lngCurMenuId = conMenu_Edit_Delete Then '报告被某人打开后再被允许它人编辑或修订
                If CheckConcurrentReport(mObjNotify.Owner, mobjStudyInfo.lngAdviceId) = False Then Exit Sub
                
                If strModuleName <> C_WORKMODULE_NAME_智能报告 Then Call UpdateReporter(mobjStudyInfo.lngAdviceId, UserInfo.姓名)
                
            End If
            
            
            Select Case strModuleName
                Case C_WORKMODULE_NAME_老版报告
                    
                    If lngCurMenuId = conMenu_PacsReport_Open Then
                        Call mobjPacsReport.PromptSave  '打开独立书写窗口之前检查嵌入式窗口的报告是否输入和保存
                        Call OpenPacsReport
                        
                        Call mObjNotify.Broadcast(BM_REPORT_EVENT_OPEN, "1", mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo)
                    Else
                        Set objCbrControl = mObjActiveMenuBar.FindControl(, lngCurMenuId + M_LNG_MENU_BASE_ID, , True)
                        
                        blnNewMenu = False
                        If objCbrControl Is Nothing Then
                            '创建隐藏的菜单项
                            Set objCbrControl = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlButton, lngCurMenuId + M_LNG_MENU_BASE_ID, "书写报告")
                            objCbrControl.Visible = False
                            blnNewMenu = True
                        End If
                        
                        If Not objCbrControl Is Nothing Then
                            lngIconId = objCbrControl.iconid
                            objCbrControl.ID = lngCurMenuId
                             'TASK:需要验证菜单ID是否为减去baseid后的值
                            Call mobjPacsReport.zlMenu.zlExecuteMenu("", lngCurMenuId)
                            
                            objCbrControl.ID = lngCurMenuId + M_LNG_MENU_BASE_ID
                            objCbrControl.iconid = lngIconId
                        End If
                        
                        If blnNewMenu Then Call objCbrControl.Delete
                    End If
                Case C_WORKMODULE_NAME_病历报告
                    Set objCbrControl = mObjActiveMenuBar.FindControl(, lngCurMenuId + M_LNG_MENU_BASE_ID, , True)
                    
                    blnNewMenu = False
                    If objCbrControl Is Nothing Then
                        '创建隐藏的菜单项
                        Set objCbrControl = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlButton, lngCurMenuId + M_LNG_MENU_BASE_ID, "书写报告")
                        objCbrControl.Visible = False
                        blnNewMenu = True
                    End If
                        
                    If Not objCbrControl Is Nothing Then
                        
                        lngIconId = objCbrControl.iconid
                        objCbrControl.ID = lngCurMenuId
            
                        
                        Call mobjEmrReport.zlExecuteCommandBars(objCbrControl)
                        
                        objCbrControl.ID = lngCurMenuId + M_LNG_MENU_BASE_ID
                        objCbrControl.iconid = lngIconId
                    End If
                    
                    If blnNewMenu Then Call objCbrControl.Delete
                Case Else
                    'PACS智能文档编辑器相关处理,仅支持查阅菜单项
                    If lngCurMenuId = conMenu_File_Open Then
                        Call PreviewRichReport(mObjNotify, mlngModule, mlngDeptID, mstrPrivs, mobjStudyInfo)
                    End If
            End Select
            
    End Select
End Sub

Public Sub IWorkMenuV2_zlUpdateMenu(ByVal strModuleName As String, Control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim lngIconId As Long
    Dim lngMenuId As Long
    
    Control.Enabled = True
    
    If Control.ID = conMenu_EditPopup + M_LNG_MENU_BASE_ID Then
        Exit Sub
    End If
    
    If mobjStudyInfo Is Nothing Then
        Control.Enabled = False
        Exit Sub
    End If
 
    lngMenuId = Control.ID - M_LNG_MENU_BASE_ID
    lngIconId = Control.iconid
    
    Control.ID = lngMenuId
    
    Select Case strModuleName
        Case C_WORKMODULE_NAME_老版报告
            Call mobjPacsReport.zlMenu.zlUpdateMenu("", Control)
        Case C_WORKMODULE_NAME_病历报告
            If mobjStudyInfo.intStep > -1 And mobjStudyInfo.intStep <= 1 Then
                Control.Enabled = False
            Else
                If Not mobjEmrReport.zlGetForm().Visible And lngMenuId <> conMenu_Edit_Modify Then
                    Control.Enabled = False
                Else
                    Call mobjEmrReport.zlUpdateCommandBars(Control)
                    
                    '如果检查已经完成，则报告的按钮将不能使用
                    If mobjStudyInfo.intStep = 6 Then
                        If InStr(Control.Caption, "查阅") > 0 Or _
                            InStr(Control.Caption, "报告检索") > 0 Or _
                            InStr(Control.Caption, "导出") > 0 Or _
                            InStr(Control.Caption, "预览") > 0 Or _
                            InStr(Control.Caption, "打印") > 0 Or _
                            InStr(Control.Caption, "复制") > 0 Or _
                            InStr(Control.Caption, "批量打印") > 0 Then
                            '如果是查阅，报告检索，或导出xml，则不修改control的enable属性
                        Else
                            Control.Enabled = False
                        End If
                    Else
                        If Control.ID = conMenu_File_Preview Or Control.ID = conMenu_File_Print Then
                            '增加是否审核后才能进行报告打印的判断
                            Control.Enabled = Control.Enabled And mobjStudyInfo.blnCanPrint
                        End If
                    End If
                End If
            End If
        Case Else
            'PACS智能文档编辑器由于没有创建相关项，因此不需要更新
    End Select
    
    '菜单补充判断
    Select Case Control.ID
        Case conMenu_PacsReport_Open
            Control.Visible = True
            
            If mobjStudyInfo.intStep = 6 Then
                Control.Enabled = CheckPopedom(mstrPrivs, "补录报告") Or InStr(Control.Caption, "查阅") > 0
            Else
                '这个补充判断，只对“PACS报告编辑器”和“电子病历编辑器”有用，“报告文档编辑器”不是用这个逻辑
                '报告的“书写”按钮，现在“PACS报告编辑器”和“电子病历编辑器”的逻辑是，不管是否有“PACS报告书写”等权限，
                '都显示此按钮，可以单击打开报告窗口，但是如果没有权限，这份报告是只读状态的。
                '没有权限也能用“书写”按钮，目的可能是嵌入式和独立窗体能看到的东西不一样，所以特地增加的这个功能？
                If strModuleName <> C_WORKMODULE_NAME_智能报告 Then
                    Control.Enabled = Not (mobjStudyInfo.intStep <= 1 And mobjStudyInfo.intStep <> -1)
                End If
            End If
            
        Case conMenu_File_Print, conMenu_File_Preview, 8352
            '补充判断平诊患者是否需要审核才能打印
'            Control.Enabled = Control.Enabled And mobjStudyInfo.blnCanPrint '增加是否审核后才能进行报告打印的判断,如果是急诊或绿色通道，允许直接打印
            
            If Control.Visible = False Or Control.Enabled = False Then
                Control.ID = lngMenuId + M_LNG_MENU_BASE_ID
                Control.iconid = lngIconId
                
                Exit Sub
            End If
            
        Case conMenu_File_NoAskPrint
            Control.Checked = mblnIsNoAskPrint
            
            
        Case conMenu_Manage_ReportFirst, conMenu_Manage_ReportSecond, _
            conMenu_Manage_ReportThird, conMenu_Manage_ReportFourth, _
            conMenu_Manage_ReportQuality

            If Not CheckPopedom(mstrPrivs, "报告质控") Or mintReportLevel = 0 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = mobjStudyInfo.intStep >= 4 And mobjStudyInfo.intStep <= 5
            End If
            
        Case conMenu_Edit_Delete
            Control.Visible = mobjStudyInfo.intStep >= 2
            
            If mobjStudyInfo.intStep = 6 Then
                '对于已完成的检查，如果当前报告未签名，且有“报告补录”权限时可以删除
                Control.Enabled = CheckPopedom(mstrPrivs, "报告补录") And mblnHasReport
            Else
                Control.Enabled = mblnHasReport And (mobjStudyInfo.intStep = 2 Or mobjStudyInfo.intStep = 3)
            End If
            
        Case conMenu_Tool_Search
            Control.Visible = True
            Control.Enabled = True
    End Select
    
    
    Control.ID = lngMenuId + M_LNG_MENU_BASE_ID
    Control.iconid = lngIconId
End Sub


Public Sub IWorkMenuV2_zlPopupMenu(ByVal strModuleName As String, objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单
    Dim objControl As CommandBarControl
    Dim objMenuControl As CommandBarControl
    
    For Each objMenuControl In mObjActiveMenuBar.ActiveMenuBar.Controls
        If objMenuControl.ID = conMenu_EditPopup + M_LNG_MENU_BASE_ID And objMenuControl.type = xtpControlPopup Then
            For Each objControl In objMenuControl.CommandBar.Controls
                If objControl.ID = conMenu_PacsReport_Open + M_LNG_MENU_BASE_ID Or _
                   objControl.ID = conMenu_Edit_Modify + M_LNG_MENU_BASE_ID Or _
                   objControl.ID = conMenu_Edit_Delete + M_LNG_MENU_BASE_ID Then
                   
                    objControl.Copy objPopup
                    
                End If
            Next
        End If
    Next
End Sub

Public Sub IWorkMenuV2_zlRefreshSubMenu(ByVal strModuleName As String, objCommandBar As Object)
'刷新弹出的子菜单
    Exit Sub
End Sub
'*********************************************************************************************

Public Sub ReplacePopupMenu(ByVal objControl As XtremeCommandBars.ICommandBarControl)
'替换右键菜单
    Dim lngOldIconId As Long
    
    '替换电子病历编辑器页面的右键菜单id
    Select Case objControl.ID
        '需要单独处理
        Case conMenu_Edit_Modify, conMenu_Edit_Audit, conMenu_Edit_Copy, _
            conMenu_Edit_Delete, conMenu_File_Open, conMenu_File_ExportToXML, _
            conMenu_Tool_Search, conMenu_Tool_SignVerify
            
            If InStr(";书写(&M);修订(&U);复制(&C);删除(&D);查阅(&R);导出XML(&L)…;报告检索(&S)…;验证签名(&V);", objControl.Caption) > 0 Then
                lngOldIconId = objControl.iconid
                
                objControl.ID = objControl.ID + M_LNG_MENU_BASE_ID
                objControl.Category = M_STR_MODULE_MENU_TAG
                
                objControl.iconid = lngOldIconId
            End If
        Case conMenu_Manage_ReportFirst, conMenu_Manage_ReportSecond, conMenu_Manage_ReportThird, conMenu_Manage_ReportFourth, conMenu_Manage_ReportQuality
            '更换报告质量的菜单id，以便在update菜单时，能够统一处理
            lngOldIconId = objControl.iconid

            objControl.ID = objControl.ID + M_LNG_MENU_BASE_ID
            objControl.Category = M_STR_MODULE_MENU_TAG

            objControl.iconid = lngOldIconId
    End Select
End Sub

Public Sub zlInitModule(objNotify As IEventNotify, _
    ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, _
    objCapLinker As Object, objMainPacsHelper As Object)
'初始化模块参数
    Dim objNewMenu As CommandBarControl
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngDeptID = lngDepartId
    mblnIsShowing = False
    mblnHasReport = False
    
    Set mobjCapLinker = objCapLinker
    Set mobjMainHelper = objMainPacsHelper
    Set mObjNotify = objNotify
    
    Call InitReportParameter

    
    '删除已经存在的报告菜单
    If Not mObjActiveMenuBar Is Nothing Then
        Set objNewMenu = mObjActiveMenuBar.FindControl(, conMenu_EditPopup, , True)
        Call DelModuleMenuAndToolBar(mObjActiveMenuBar, objNewMenu, M_STR_MODULE_MENU_TAG)
    End If

End Sub


Private Sub DisposeCurReportObj()
'释放当前的报告编辑器对象

    If Not mobjPacsReport Is Nothing Then
        Call mobjPacsReport.SetMenuDownState(False)
        Unload mobjPacsReport
        Set mobjPacsReport = Nothing
    End If
    
    If Not mobjEmrReport Is Nothing Then
        Unload mobjEmrReport.zlGetForm
        Set mobjEmrReport = Nothing
    End If
    
    Set mobjMainHelper = Nothing
End Sub


Public Sub InitReportParameter()
'初始化报告参数
'    mReportType = GetDeptPara(mlngDeptID, "报告编辑器", 0)                '报告编辑器
    mblnIsNoAskPrint = zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1"
    
    mblnTechReptSame = False
    mblnIsEditWithRepImage = False
    
    '病理系统不需要判断检查技师是否与登录用户相同
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        mblnTechReptSame = GetDeptPara(mlngDeptID, "只能填写自己检查的报告", 0) = "1"  '只能填写自己检查的报告
    End If
    
    mblnIsEditWithRepImage = Val(GetDeptPara(mlngDeptID, "有图像才能写报告", 0)) = 1
    
    
    mstrReportLevel = nvl(GetDeptPara(mlngDeptID, "报告质量等级", "甲,乙"))
    mintReportLevel = Val(GetDeptPara(mlngDeptID, "报告质量判定", "0"))
End Sub

Public Sub AllowLocate(blnIsAllowLocate As Boolean)
'    If mReportType = PACS报告编辑器 Then
'        Call mfrmPacsReport.AllowLocate(blnIsAllowLocate)
'    End If
End Sub


Private Function HasReport(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean) As Boolean
 
'更新是否存在报告
    Dim rsTemp As Recordset
    Dim strSQL As String

    strSQL = "Select 病历ID From 病人医嘱报告 Where 医嘱ID= [1]"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否存在报告", lngAdviceId)
    
    If rsTemp.RecordCount < 1 Then
        HasReport = False
        Exit Function
    End If
         
    HasReport = Len(nvl(rsTemp!病历Id)) > 0
 
End Function

Public Sub zlRefreshFace(objStudyInfo As clsStudyInfo, Optional ByVal strModuleName As String = "", _
            Optional ByVal blnIsForceRefresh As Boolean = False, Optional ByVal blnIsHistory As Boolean = False)
'刷新界面
On Error GoTo errhandle
    
    If Not mobjStudyInfo Is Nothing And Not objStudyInfo Is Nothing Then
        '如果检查信息相同，且不是强制刷新则退出
        If mblnIsRefreshStudy = True Then   '已经刷新过则需要判断检查信息相同以及是否进行强制刷新
            If mobjStudyInfo.IsEquals(objStudyInfo) And blnIsForceRefresh = False Then Exit Sub
        Else
            '如果检查信息没有被刷新过，则进行强制刷新
            blnIsForceRefresh = True
        End If
    End If
    
    mblnHasReport = False
    mblnIsProcessing = True
    
    Set mobjStudyInfo = objStudyInfo
    
    Select Case strModuleName
        Case C_WORKMODULE_NAME_老版报告
            mblnHasReport = HasReport(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
            
            Call mobjPacsReport.zlRefresh(mobjStudyInfo, blnIsForceRefresh, blnIsHistory)
    
        Case C_WORKMODULE_NAME_病历报告
            mblnHasReport = HasReport(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)

            Call mobjEmrReport.zlRefresh(mobjStudyInfo.lngAdviceId, mlngDeptID, IIf(blnIsHistory, False, True), mobjStudyInfo.blnMoved, mobjStudyInfo.blnCanPrint, mlngModule)
            
        Case Else
            'PACS智能文档编辑器相关处理
            mblnHasReport = False
            Call mobjRichReportView.zlRefresh(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved, blnIsForceRefresh)
    End Select
        
    mblnIsRefreshStudy = True
    mblnIsProcessing = False
    Exit Sub
errhandle:
    mblnIsProcessing = False
    If ErrCenter() = 1 Then Resume
End Sub
 
Public Sub RefreshPopupWindow(objStudyInfo As clsStudyInfo)
'刷新弹出报告窗口
    Dim objForm As Object
    
    '查找窗口数组,找到需要打开的窗体，则通过Zorder把窗体显示到最前面
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.AdviceId = objStudyInfo.lngAdviceId And objForm.IsLinkHelper = False Then
                Call objForm.zlRefresh(objStudyInfo)
                
                Exit Sub
            End If
        End If
    Next
End Sub

Public Sub OpenPacsReport()
On Error GoTo errhandle
    Dim i As Integer
    Dim j As Integer
    Dim strDkpTitle As String
    Dim objForm As Object
    
    Dim objReport As frmReportV2
    
    '如果没有选择行数据，则直接退出执行
    If mobjStudyInfo Is Nothing Then
        MsgBoxD mObjNotify.Owner, M_STR_HINT_NoSelectData, vbInformation, "提示"
        Exit Sub
    End If
    
    '判断是否存在已经打开的报告编辑框
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.AdviceId = mobjStudyInfo.lngAdviceId And objForm.IsLinkHelper = False Then
                objForm.WindowState = 0
                objForm.Visible = True
                objForm.ZOrder
                
                Call objForm.Shake
                
                Exit Sub
            End If
        End If
    Next
    

    Set objReport = New frmReportV2
    
    objReport.zlInit mObjNotify, mlngModule, mlngDeptID, mstrPrivs, mobjCapLinker, Nothing
    objReport.zlRefresh mobjStudyInfo
    
    Call objReport.Show(0, mObjNotify.Owner)
    Call objReport.SetReportTitle(mobjStudyInfo)
    Call objReport.ReSetFormFontSize(mobjPacsReport.FontSize)
    
    '弹出式方式进行报告编辑时，需要将焦点定位到编辑框
    objReport.LocateEditBox
    
    Exit Sub
errhandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub BatPrint()
On Error GoTo errhandle
    Dim objPacsReport As frmReportV2
    Dim objEprWarp As frmEPREditWrapV2
    
    Dim strReturn As String
    Dim i As Integer
    Dim strReportString As String
    Dim lngAdviceId As Long
    Dim blnIsMoved As Boolean
    Dim objBatStudyInfo As clsStudyInfo
    
    '选病人
    strReturn = ""
    If Len(strReturn) <= 0 Then Exit Sub
    
    '循环调用报告打印
    '对于使用PACS报告编辑器打印的，调用PACS报告编辑器窗口来打印
    '返回值由"医嘱ID-是否PACS报告编辑器-执行科室ID|医嘱ID-是否PACS报告编辑器-执行科室ID|..."组成
    For i = 0 To UBound(Split(strReturn, "|"))
        strReportString = Split(strReturn, "|")(i)
        
        lngAdviceId = CLng(Split(strReportString, "-")(0))
        blnIsMoved = IIf((Val(Split(strReportString, "-")(3))) = 2, True, False)

        Set objBatStudyInfo = zlGetStudyAdvice(lngAdviceId)
        objBatStudyInfo.blnCanPrint = True
            
        '判断是否使用PACS报告编辑器
        If Split(strReportString, "-")(1) = 1 Then  '使用PACS报告编辑器
            If objPacsReport Is Nothing Then
                Set objPacsReport = New frmReportV2
                Call objPacsReport.zlInit(mObjNotify, mlngModule, mlngDeptID, mstrPrivs, mobjCapLinker, mobjMainHelper)
            End If
            
            Call objPacsReport.zlRefresh(objBatStudyInfo)
            Call objPacsReport.zlMenu.zlExecuteMenu("", conMenu_File_Print)
            
        ElseIf Split(strReportString, "-")(1) = 0 Then    '使用病历编辑器

            If objEprWarp Is Nothing Then
                Set objEprWarp = New frmEPREditWrapV2
                Call objEprWarp.InitEprEditor(Nothing, mObjNotify, mlngModule, mlngDeptID)
            End If
            
            
            Call objEprWarp.ExecuteMenu(objBatStudyInfo, conMenu_File_BatPrint)
        Else
            'TODO:PACS智能文档编辑器相关处理
        End If
    Next
    
    Unload objPacsReport
    
    Set objPacsReport = Nothing
    Set objEprWarp = Nothing
    Set objBatStudyInfo = Nothing
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'Public Sub SingPrint(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean, Optional ByVal strPrintFmts As String = "")
''单独打印，由主界面调用
'On Error GoTo errhandle
'    Dim objPacsReport As frmReportV2
'    Dim objEprWarp As frmEPREditWrapV2
'    Dim objRichV2 As frmRichReportV2
'
'    Dim objSignStudyInfo As clsStudyInfo
'    Dim lngReportType As Long
'
'    Set objSignStudyInfo = zlGetStudyAdvice(lngAdviceId)
'
'    lngReportType = Val(GetDeptPara(objSignStudyInfo.lngExeDepartmentId, "报告编辑器", 0))                '报告编辑器类型
'
'    '判断是否使用PACS报告编辑器
'    If lngReportType = PACS报告编辑器 Then  '使用PACS报告编辑器
'        If objPacsReport Is Nothing Then
'            Set objPacsReport = New frmReportV2
'            Call objPacsReport.zlInit(mObjNotify, mlngModule, mlngDeptID, mstrPrivs, mobjCapLinker, mobjMainHelper)
'        End If
'
'        Call objPacsReport.zlRefresh(objSignStudyInfo)
'
'        If Len(strPrintFmts) > 0 Then Call objPacsReport.SetPrintFmts(Split(strPrintFmts, ":")(0), Split(strPrintFmts & ":", ":")(1))
'
'        Call objPacsReport.zlMenu.zlExecuteMenu("", conMenu_File_Print)
'
'        Unload objPacsReport
'
'    ElseIf lngReportType = 电子病历编辑器 Then   '使用病历编辑器
'
'        Set objEprWarp = New frmEPREditWrapV2
'        Call objEprWarp.InitEprEditor(Nothing, mObjNotify, mlngModule, mlngDeptID)
'
'        Call objEprWarp.ExecuteMenu(objSignStudyInfo, conMenu_File_BatPrint)
'
'        Unload objEprWarp    'TODO:判断是否会造成异常退出
'    Else
'        Set objRichV2 = New frmRichReportV2
'
'        Call objRichV2.zlInit(Me, mlngModule, mlngDeptID, mstrPrivs)
'
'        Call objRichV2.zlRefresh(objSignStudyInfo.lngAdviceId, objSignStudyInfo.blnMoved, True)
'
'        Call objRichV2.PrintPreview(False)
'
'        Unload objRichV2
'
'    End If
'
'
'    Set objPacsReport = Nothing
'    Set objEprWarp = Nothing
'    Set objRichV2 = Nothing
'    Set objSignStudyInfo = Nothing
'Exit Sub
'errhandle:
'    If ErrCenter() = 1 Then Resume
'End Sub


Private Function SetReportQuality(ByVal lngID As Long, ByVal strReportLevel As String) As Long
    On Error GoTo errhandle
    Dim strSQL As String
    Dim strResult As String
    Dim strGrades() As String

    SetReportQuality = 0
    
    If mobjStudyInfo Is Nothing Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, gstrSysName
        Exit Function
    End If
    
    Select Case lngID
        Case conMenu_Manage_ReportFirst
            strResult = 1
        Case conMenu_Manage_ReportSecond
            strResult = 2
        Case conMenu_Manage_ReportThird
            strResult = 3
        Case conMenu_Manage_ReportFourth
            strResult = 4
    End Select

    strSQL = "Zl_报告质量_Update(" & mobjStudyInfo.lngAdviceId & ",'" & strResult & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "报告质量")

    SetReportQuality = Val(strResult)

Exit Function
errhandle:
    If ErrCenter = 1 Then Resume
End Function


Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
    If Not mobjPacsReport Is Nothing Then
        Call mobjPacsReport.ReSetFormFontSize(bytFontSize)
    End If
    
End Sub


Private Sub mobjEmrReport_AfterSaved(ByVal lngOrderID As Long, ByVal lngSaveType As Long)
'发送报告保存消息
'注：回退操作会触发此保存事件
    Select Case lngSaveType
        Case 0
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_SAVE, 1, lngOrderID, 0, -1)   '-1表示病历编辑器
            mblnHasReport = True
        Case 1
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_SIGN, 1, lngOrderID, 0, -1)
            mblnHasReport = True
        Case 2
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_AUDIT, 1, lngOrderID, 0, -1)
            mblnHasReport = True
    End Select
End Sub
 

Private Sub mobjEmrReport_AfterClosed(ByVal lngOrderID As Long)
'刷新嵌入报告内容
    Call mobjEmrReport.zlRefresh(mobjStudyInfo.lngAdviceId, mlngDeptID, True, mobjStudyInfo.blnMoved, mobjStudyInfo.blnCanPrint, mlngModule, True)
    
    Call UpdateReporter(lngOrderID, "")
End Sub
 
 
Private Sub mobjEmrReport_AfterDeleted(ByVal lngOrderID As Long)
'发送报告删除消息
    Call mObjNotify.SendRequest(WM_LIST_SYNCROW, 1, lngOrderID)
    
    mblnHasReport = False
End Sub
  

Private Sub mobjEmrReport_AfterPrinted(ByVal lngOrderID As Long)
'发送报告打印消息
    Call mObjNotify.Broadcast(BM_REPORT_EVENT_PRINT, 1, lngOrderID, 0, -1)
End Sub

   

Public Sub DoFontSize(ByVal blIsDock As Boolean, ByVal intFontSize As Integer)
'同步字号处理：独立窗体改变报告字号同时改变嵌入式窗体字号。
    If Not blIsDock Then Exit Sub
    
    Call mobjPacsReport.SetMeneFontSize(intFontSize)
End Sub
  
