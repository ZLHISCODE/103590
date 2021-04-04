VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkModule_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IWorkMenu


Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "报告"
Private Const M_LNG_MENU_BASE_ID As Long = 1000000

Private mlngModule  As Long
Private mstrPrivs  As String
Private mlngDeptID As Long
Private mobjOwner As Object

Private mlngAdviceID As Long
Private mlngSendNo As Long
Private mblnMoved As Boolean
Private mlngStudyState As Long

Private mblnReadOnly As Boolean

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mObjActiveMenuBar As CommandBars
Private mblnRefreshState As Boolean

Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer     '观片站对象
Attribute mobjPacsCore.VB_VarHelpID = -1
Private WithEvents mobjReport As zlRichEPR.cDockReport  '电子病历报告对象
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mfrmPacsReport As frmReport     'PACS报告编辑器，嵌入主程序的窗口
Attribute mfrmPacsReport.VB_VarHelpID = -1
Private WithEvents mfrmPacsReportDock As frmReport      'PACS报告编辑器,独立窗口
Attribute mfrmPacsReportDock.VB_VarHelpID = -1
Private mobjPacsReportArry() As frmReport                   'PACS报告编辑器数组

Private mobjReportContainer As Object
Private mobjStudyList As ucFlexGrid
Private mblnIsShowing As Boolean


'报告相关参数
Private mblnTechReptSame As Boolean
Private mblnIsPacsReport As Boolean
Private mblnIsShowVideoCapture As Boolean
Private mblnShowImgAtReport As Boolean
Private mblnLocalizerBackward As Boolean

Private mblnCanPrint As Boolean
Private mblnIsHistory As Boolean
Private mstrDoDoctor As String  '检查技师
Private mstrStudyUID As String '检查UID

Private mstrReportLevel As String    '报告质量等级串

'本窗体的事件
Public Event AfterOpen()
Public Event BeforeEdit(ByVal lngOrderID As Long)

'frmOwnerForm主要用于在该事件执行时,如果在该事件中有模态窗口显示，就需要制定owner拥有者，且必须使用该参数，否则会
'造成程序处于假死状态，如在电子病例签名后，选择阳性率，阳性率窗口在未设置该参数的情况下，将造成这种情况。
Public Event AfterSaved(ByVal lngOrderID As Long, frmOwnerForm As Object, ByVal lngSaveType As Long)
Public Event AfterClosed(ByVal lngOrderID As Long)
Public Event AfterPrinted(ByVal lngOrderID As Long)
Public Event AfterDeleted(ByVal lngOrderID As Long)



'获取菜单接口对象
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property

'获取报告窗口句柄
Property Get zlGetForm() As Object
    If mblnIsPacsReport Then
        Set zlGetForm = mfrmPacsReport
    Else
        Set zlGetForm = mobjReport.zlGetForm()
    End If
End Property


'设置报告的图像处理对象
Property Get PacsCore() As zl9PacsCore.clsViewer
    Set PacsCore = mobjPacsCore
End Property

Property Set PacsCore(objPacsCore As zl9PacsCore.clsViewer)
    Set mobjPacsCore = objPacsCore
    Set mfrmPacsReport.PacsCore = objPacsCore
    Set mobjReport.PacsCore = objPacsCore
End Property


Public Sub NotificationRefresh()
'通知刷新
    mblnRefreshState = False
    
    If mblnIsPacsReport Then Call mfrmPacsReport.NotificationRefresh
End Sub


Private Sub Class_Initialize()
    mblnIsPacsReport = True
    
    '在这里创建所需的所有报告对象，避免重复创建pacs报告编辑器时，因所嵌入的视频采集对象的parent被释放带来的异常
    Set mobjReport = New zlRichEPR.cDockReport
    
    Set mfrmPacsReport = New frmReport
    
    SetWindowStyle mfrmPacsReport.hWnd
End Sub

Public Sub Free()
On Error GoTo errHandle
    Dim i As Long
    
    If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
        '卸载打开的独立报告编辑器窗口
        For i = LBound(mobjPacsReportArry) To UBound(mobjPacsReportArry)
            If Not mobjPacsReportArry(i) Is Nothing Then
                Unload mobjPacsReportArry(i)
                Set mobjPacsReportArry(i) = Nothing
            End If
        Next i
    End If
    
    If Not mfrmPacsReportDock Is Nothing Then
        Unload mfrmPacsReportDock
        Set mfrmPacsReportDock = Nothing
    End If
    
    Call DisposeCurReportObj
    
errHandle:
End Sub



'接口实现部分*********************************************************************************

Public Function IWorkMenu_zlIsModuleMenu(ByVal objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单
'医嘱，费用，病历模块创建的菜单名对应为“医嘱”，“费用”，“病历”
    
    IWorkMenu_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenu_zlCreateMenu(objMenuBar As Object)
'创建影像记录对应的菜单
    Dim objNewMenu As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    
    Set mObjActiveMenuBar = objMenuBar
    
    Set objNewMenu = objMenuBar.FindControl(, conMenu_EditPopup)

    '如果菜单存在，则先删除对应的菜单和工具栏按钮重新创建
    Call DelModuleMenuAndToolBar(objMenuBar, objNewMenu, M_STR_MODULE_MENU_TAG)
        
    If mblnIsPacsReport Then
        Call mfrmPacsReport.zlMenu.zlCreateMenu(objMenuBar)
        Call mfrmPacsReport.zlMenu.zlCreateToolBar(objMenuBar.Item(2))
    Else
        Call mobjReport.zlDefCommandBars(mObjActiveMenuBar)
        '添加预览、打印、批量打印的菜单栏和工具栏按钮
        Call zlDefCommandBars(mObjActiveMenuBar)
    End If
    
    '创建报告质量菜单
    Set cbrMenuBar = mObjActiveMenuBar.FindControl(, conMenu_ManagePopup)
    With cbrMenuBar.CommandBar
        If InStr(mstrPrivs, "报告质控") > 0 Then
            Set objNewMenu = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_ReportQuality, "报告质量", "", 8346, False, cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_GChannel).Index - 1)
            objNewMenu.Category = ""
            Call CreateSubordinateMenuTools(mstrReportLevel, objNewMenu)
        End If
    End With
    
    Call SetMenuCategory(objMenuBar, objNewMenu, M_STR_MODULE_MENU_TAG)
    '避免创建模块工具栏时，窗口左上角出现不完全显示的图标，一闪而过的现象
    Call objMenuBar.RecalcLayout
End Sub

Private Sub zlDefCommandBars(objMenuBar As Object)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    
    Set cbrMenuBar = objMenuBar.ActiveMenuBar.Controls.Find(, conMenu_EditPopup)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Preview, "预览", "", 102, False, 1)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Print, "打印", "", 103, False, 2)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_BatPrint, "批量打印", "", 122, False, 3)
        If cbrMenuBar.CommandBar.Controls.Count >= 4 Then cbrMenuBar.CommandBar.Controls(4).BeginGroup = True
    End With
    
    Set cbrToolBar = objMenuBar(2)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Preview, "预览", "", 102, True, cbrControl.Index + 1)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Print, "打印", "", 103, False, cbrControl.Index + 1)
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
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
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
                    
                    lngIconId = objSubMenu.Controls(j).IconId
                    
                    objSubMenu.Controls(j).ID = objSubMenu.Controls(j).ID + M_LNG_MENU_BASE_ID
                    objSubMenu.Controls(j).IconId = lngIconId
                End If
            Next j
            
            If objActiveMenu.Controls(i).Category = "" Then
                objActiveMenu.Controls(i).Category = strCategory
                
                lngIconId = objActiveMenu.Controls(i).IconId
                
                objActiveMenu.Controls(i).ID = objActiveMenu.Controls(i).ID + M_LNG_MENU_BASE_ID
                objActiveMenu.Controls(i).IconId = lngIconId
            End If
        End If
    Next i
    
    Set objToolBar = objMenuBar.Item(2)
    If Not objToolBar Is Nothing Then
        For i = 1 To objToolBar.Controls.Count
            If objToolBar.Controls(i).Category = "" Then
                objToolBar.Controls(i).Category = strCategory
                
                '避免因修改id造成图标不能显示
                lngIconId = objToolBar.Controls(i).IconId
                
                objToolBar.Controls(i).ID = objToolBar.Controls(i).ID + M_LNG_MENU_BASE_ID
                objToolBar.Controls(i).IconId = lngIconId
            End If
            
            If objToolBar.Controls(i).type <> xtpControlButton And objToolBar.Controls(i).type <> xtpControlButtonPopup Then
                Set objSubMenu = objToolBar.Controls(i).CommandBar
                
                For j = objSubMenu.Controls.Count To 1 Step -1
                    If objSubMenu.Controls(j).Category = "" Then
                        objSubMenu.Controls(j).Category = strCategory
                        
                        lngIconId = objSubMenu.Controls(j).IconId
                        
                        objSubMenu.Controls(j).ID = objSubMenu.Controls(j).ID + M_LNG_MENU_BASE_ID
                        objSubMenu.Controls(j).IconId = lngIconId
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


Public Sub IWorkMenu_zlCreateToolBar(objToolBar As Object)
'创建工具栏
    Dim cbrControl As CommandBarControl

    '创建报告质量工具栏
    If InStr(mstrPrivs, "报告质控") > 0 Then
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_ReportQuality, "报告质量", "报告质量", 8346, False, IIf(mlngModule <> G_LNG_PATHOLSYS_NUM, mObjActiveMenuBar.FindControl(, conMenu_Manage_Complete).Index - 1, mObjActiveMenuBar.FindControl(, conMenu_Manage_Complete).Index))
        cbrControl.Category = "" 'M_STR_MODULE_MENU_TAG
        Call CreateSubordinateMenuTools(mstrReportLevel, cbrControl)
    End If
    
    Call SetMenuCategory(mObjActiveMenuBar, Nothing, M_STR_MODULE_MENU_TAG)
    
    Exit Sub
'    If mblnIsPacsReport Then Call mfrmPacsReport.zlMenu.zlCreateToolBar(objToolBar)
End Sub
    
Public Sub IWorkMenu_zlClearMenu()
'清除所创建的菜单
    Dim cbrControl As CommandBarControl

    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_EditPopup + M_LNG_MENU_BASE_ID)
'    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    Call DelModuleMenuAndToolBar(mObjActiveMenuBar, cbrControl, M_STR_MODULE_MENU_TAG)
End Sub


Public Sub IWorkMenu_zlClearToolBar()
'清除创建的工具栏

    '菜单删除时，已同步删除了对应的工具栏
    Exit Sub
End Sub

Public Sub IWorkMenu_zlExecuteMenu(ByVal lngMenuId As Long)
'根据菜单ID执行对应功能
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    Dim lngCurMenuId As Long
    Dim lngIconId As Long
    
    lngCurMenuId = lngMenuId - M_LNG_MENU_BASE_ID
    
    If lngCurMenuId = conMenu_File_BatPrint Then
        Call BatPrint
        Exit Sub
    End If
    
    If lngCurMenuId = conMenu_Manage_ReportFirst Or lngCurMenuId = conMenu_Manage_ReportSecond Or lngCurMenuId = conMenu_Manage_ReportThird Or lngCurMenuId = conMenu_Manage_ReportFourth Then
        Call Menu_Manage_ReportQuality(lngCurMenuId, mstrReportLevel, mobjStudyList)
        Exit Sub
    End If
    
    '报告被某人打开后再被允许它人编辑或修订
    If lngCurMenuId = conMenu_Edit_Audit Or lngCurMenuId = conMenu_Edit_Modify Or lngCurMenuId = conMenu_PacsReport_Open Or lngCurMenuId = conMenu_Edit_Delete Then
        If CheckConcurrentReport(mobjOwner, mlngAdviceID) = True Then
            Call UpdateReporter(mlngAdviceID, UserInfo.姓名)
        Else
            Exit Sub
        End If
    End If
    
    '控制 只能书写自己检查的报告,'不允许书写、修订、删除
    If mblnTechReptSame = True _
        And (lngCurMenuId = conMenu_Edit_Modify Or lngCurMenuId = conMenu_Edit_Audit Or lngCurMenuId = conMenu_PacsReport_Open Or lngCurMenuId = conMenu_Edit_Delete) _
        And mstrDoDoctor <> "" And mstrDoDoctor <> UserInfo.姓名 Then
        
        Call UpdateReporter(Val(mlngAdviceID), "")
        Call MsgBoxD(mobjOwner, "你不是这个患者的检查技师，无法操作这份报告。", vbInformation, gstrSysName)
    Else
        If mblnIsPacsReport Then
            If lngCurMenuId = conMenu_PacsReport_Open Then
                Call OpenPacsReport
            Else
                Set objCbrControl = mObjActiveMenuBar.FindControl(, lngCurMenuId + M_LNG_MENU_BASE_ID, , True)
                
                If Not objCbrControl Is Nothing Then
                    lngIconId = objCbrControl.IconId
                    objCbrControl.ID = lngCurMenuId
                    
                    Call mfrmPacsReport.zlMenu.zlExecuteMenu(lngCurMenuId)
                    
                    objCbrControl.ID = lngCurMenuId + M_LNG_MENU_BASE_ID
                    objCbrControl.IconId = lngIconId
                End If
            End If
        Else
            Set objCbrControl = mObjActiveMenuBar.FindControl(, lngCurMenuId + M_LNG_MENU_BASE_ID, , True)
            
            If Not objCbrControl Is Nothing Then
                
                lngIconId = objCbrControl.IconId
                objCbrControl.ID = lngCurMenuId
                
                Call mobjReport.zlExecuteCommandBars(objCbrControl)
                
                objCbrControl.ID = lngCurMenuId + M_LNG_MENU_BASE_ID
                objCbrControl.IconId = lngIconId
                
                '在使用电子病历编辑器的时候，勾选“书写时打开观片站”参数后，打开观片站,只在书写的情况下打开
                If lngCurMenuId = conMenu_Edit_Modify Then Call OpenViewerWithReport
            End If
        End If
    End If
    
'    '清除报告“处理中”标记之后，需要刷新病人列表
'    If control.ID = conMenu_PacsReport_ClearWritingState Then
'        Call RefreshList
'    End If
                    

    
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    Dim lngIconId As Long
    Dim lngMenuId As Long
    
    control.Enabled = True
    
    If control.ID = conMenu_EditPopup + M_LNG_MENU_BASE_ID Then
        Exit Sub
    End If
 
    lngMenuId = control.ID - M_LNG_MENU_BASE_ID
    lngIconId = control.IconId
    
    control.ID = lngMenuId
    
    If mblnIsPacsReport Then
        Call mfrmPacsReport.zlMenu.zlUpdateMenu(control)
    Else
        If mlngStudyState > -1 And mlngStudyState <= 1 Then
            control.Enabled = False
        Else
            If Not mobjReport.zlGetForm().Visible And lngMenuId <> conMenu_Edit_Modify Then
                control.Enabled = False
            Else
                Call mobjReport.zlUpdateCommandBars(control)
                
                '如果检查已经完成，则报告的按钮将不能使用
                If mlngStudyState = 6 Then
                    If InStr(control.Caption, "查阅") > 0 Or _
                        InStr(control.Caption, "报告检索") > 0 Or _
                        InStr(control.Caption, "导出") > 0 Or _
                        InStr(control.Caption, "预览") > 0 Or _
                        InStr(control.Caption, "打印") > 0 Or _
                        InStr(control.Caption, "批量打印") > 0 Then
                        '如果是查阅，报告检索，或导出xml，则不修改control的enable属性
                    Else
                        control.Enabled = False
                    End If
                End If
            End If
        End If
    End If
    
    '菜单补充判断
    Select Case control.ID

        Case conMenu_PacsReport_ClearWritingState, conMenu_PacsReport_Open
            If (mlngStudyState <= 1 And mlngStudyState <> -1) Or mlngStudyState = 6 Then control.Enabled = False

        Case conMenu_File_Open, conMenu_File_ExportToXML, conMenu_File_Print, conMenu_File_Preview
            If mlngStudyState <= 1 Then control.Enabled = False

        Case conMenu_Tool_Search

        Case conMenu_Edit_Delete
            If mlngStudyState >= 4 Or (mlngStudyState <= 1) Then control.Enabled = False    '如果检查过程大于或等于4，则不允许删除检查报告
        
        Case conMenu_Manage_ReportFirst, conMenu_Manage_ReportSecond, _
            conMenu_Manage_ReportThird, conMenu_Manage_ReportFourth, _
            conMenu_Manage_ReportQuality
            
            If InStr(mstrPrivs, "报告质控") <= 0 Then
                control.Visible = False
            Else
                control.Visible = True
                control.Enabled = mlngStudyState >= 2 And mlngStudyState <= 5
            End If
    End Select
    
    
    control.ID = lngMenuId + M_LNG_MENU_BASE_ID
    control.IconId = lngIconId
End Sub


Public Sub IWorkMenu_zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
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

Public Sub IWorkMenu_zlRefreshSubMenu(objCommandBar As Object)
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
                lngOldIconId = objControl.IconId
                
                objControl.ID = objControl.ID + M_LNG_MENU_BASE_ID
                objControl.Category = M_STR_MODULE_MENU_TAG
                
                objControl.IconId = lngOldIconId
            End If
'        Case conMenu_Manage_ReportFirst, conMenu_Manage_ReportSecond, conMenu_Manage_ReportThird, conMenu_Manage_ReportFourth, conMenu_Manage_ReportQuality
'            '更换报告质量的菜单id，以便在update菜单时，能够统一处理
'            lngOldIconId = objControl.IconId
'
'            objControl.ID = objControl.ID + M_LNG_MENU_BASE_ID
'            objControl.Category = M_STR_MODULE_MENU_TAG
'
'            objControl.IconId = lngOldIconId
    End Select
End Sub

Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing)
'初始化模块参数
    Dim objNewMenu As CommandBarControl
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngDeptID = lngDepartId
    mblnIsShowing = False
    
    If Not owner Is Nothing Then Set mobjOwner = owner
    
    Call InitReportParameter

    
    '删除已经存在的报告菜单
    If Not mObjActiveMenuBar Is Nothing Then
        Set objNewMenu = mObjActiveMenuBar.FindControl(, conMenu_EditPopup, , True)
        Call DelModuleMenuAndToolBar(mObjActiveMenuBar, objNewMenu, M_STR_MODULE_MENU_TAG)
    End If
    
    '创建所需的报告编辑器对象
    If mblnIsPacsReport Then
        Call mfrmPacsReport.zlInitModule(lngModule, strPrivs, lngDepartId, owner)
    End If
End Sub


Private Sub DisposeCurReportObj()
'释放当前的报告编辑器对象

    If Not mfrmPacsReport Is Nothing Then
        Unload mfrmPacsReport
        Set mfrmPacsReport = Nothing
    End If
    
    If Not mobjReport Is Nothing Then
        Unload mobjReport.zlGetForm
        Set mobjReport = Nothing
    End If
End Sub


Public Sub InitReportParameter()
'初始化报告参数
    mblnIsPacsReport = GetDeptPara(mlngDeptID, "报告编辑器", 0) = "1" '              '报告编辑器
    
    mblnTechReptSame = False
    
    '病理系统不需要判断检查技师是否与登录用户相同
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        mblnTechReptSame = GetDeptPara(mlngDeptID, "只能填写自己检查的报告", 0) = "1"  '只能填写自己检查的报告
    End If
    
    mblnIsShowVideoCapture = GetDeptPara(mlngDeptID, "显示视频采集", 0) = "1" '显示视频采集
    
    mblnShowImgAtReport = (Val(zlDatabase.GetPara("报告时观片", glngSys, mlngModule, 0)) = 1)    '报告时观片
    mblnLocalizerBackward = GetDeptPara(mlngDeptID, "定位片后置", 0) = "1" '       '定位片后置
    
    mstrReportLevel = Nvl(GetDeptPara(mlngDeptID, "报告质量等级", "甲,乙"))
End Sub


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
    
    If Not mfrmPacsReport Is Nothing Then Call mfrmPacsReport.zlUpdateAdviceInf(lngAdviceID, lngSendNO, lngStudyState, blnMoved)
End Sub


Public Sub zlUpdateOtherInf(objReportContainer As Object, objStudyList As Object, _
    ByVal blnIsHistory As Boolean, ByVal blnCanPrint As Boolean, ByVal strDoDoctor As String, ByVal strStudyUID As String)
'更新报告需要的其他信息
    Set mobjReportContainer = objReportContainer
    Set mobjStudyList = objStudyList
    
    mblnIsHistory = blnIsHistory
    mblnCanPrint = blnCanPrint
    mstrDoDoctor = strDoDoctor
    mstrStudyUID = strStudyUID
End Sub


Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False, Optional blnShowVideo As Boolean = True, Optional blnShowReport As Boolean = True)
'刷新界面
On Error GoTo errHandle

    If blnShowReport Then
        '将报告窗口嵌入到主界面的容器中
        Call RefreshReportWindow
    End If
    

    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo _
        And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
        
    mlngTmpAdviceId = mlngAdviceID
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True
    
    
    If mblnIsPacsReport Then
        Call mfrmPacsReport.zlRefreshFace(blnForceRefresh)
        
        If GetActiveWindow = mobjOwner.hWnd And blnShowVideo Then Call mfrmPacsReport.RefreshVideo
    Else
        Call mobjReport.zlRefresh(mlngAdviceID, mlngDeptID, Not mblnIsHistory, mblnMoved, mblnCanPrint, mlngModule)
    End If
    
    Call UpdateSize
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub zlShowReportVideo()
'显示报告视频
    If mblnIsPacsReport Then mfrmPacsReport.RefreshVideo
End Sub

'刷新报告窗口
Private Sub RefreshReportWindow()
    Dim lngHandle As Long
    Dim lngStyle As Long
    
    If mobjReportContainer Is Nothing Then Exit Sub
    
    If Not mblnIsShowing Then
        If mblnIsPacsReport Then
            Call mfrmPacsReport.Show
        Else
            Call mobjReport.zlGetForm().Show
        End If
        
        mblnIsShowing = True
    End If
    
    If mblnIsPacsReport Then
        lngHandle = mfrmPacsReport.hWnd
        
'        mobjReport.zlGetForm.Visible = False
'        SetParent mobjReport.zlGetForm.hWnd, 0
    Else
        lngHandle = mobjReport.zlGetForm().hWnd
        
'        mfrmPacsReport.Visible = False
'        SetParent mfrmPacsReport.hWnd, 0
    End If
    
    Call ShowSubWindow(lngHandle, mobjReportContainer.hWnd)
    '必须在这里设置一次窗口样式，否则可能造成嵌入的报告窗口包含标题栏
    Call SetWindowStyle(lngHandle)
    
    If mobjReportContainer.Tag = lngHandle Then Exit Sub
    
'    Call SetParent(lngHandle, mobjReportContainer.hWnd)
    
    mobjReportContainer.Tag = lngHandle
    If mlngAdviceID > 0 Then mobjReportContainer.Visible = True
    
    Call UpdateSize
End Sub

Private Sub SetWindowStyle(ByVal lngHandle As Long)
    Dim lngWindowStyle As Long
    
    lngWindowStyle = GetWindowLong(lngHandle, GWL_STYLE)
    
'    If (lngWindowStyle And WS_CHILD) = WS_CHILD Then Exit Sub
    
    lngWindowStyle = lngWindowStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)

    Call SetWindowLong(lngHandle, GWL_STYLE, lngWindowStyle Or WS_CHILD)
End Sub

Public Sub ShowSubWindow(ByVal lngHwnd As Long, Optional ByVal lngMainHwnd As Long)
'功能：显示指定的窗体，以子窗体方式
'参数：lngHwnd=要作为子窗体显示的窗体的句柄
'      lngMainHwnd=父窗体句柄，不传时表明不以它的子窗体显示
'说明：该项函数主要用于在ZLBH中融合调用ZLHIS窗体显示
    Dim vRect1 As RECT, vRect2 As RECT
    Dim X As Long, Y As Long
    
    If lngHwnd <= 0 Then Exit Sub
    
    If lngMainHwnd <> 0 Then
        SetParent lngHwnd, lngMainHwnd
    Else
        SetParent lngHwnd, 0
    End If

    '显示在父窗体的中央
    If IsWindowVisible(lngHwnd) = 0 Then
        GetWindowRect lngHwnd, vRect1
        GetWindowRect lngMainHwnd, vRect2
        
        X = ((vRect2.Right - vRect2.Left) - (vRect1.Right - vRect1.Left)) / 2
        Y = ((vRect2.Bottom - vRect2.Top) - (vRect1.Bottom - vRect1.Top)) / 2
        If X < 0 Then X = 0
        If Y < 0 Then Y = 0
        
        SetWindowPos lngHwnd, 0, X, Y, 0, 0, &H40 Or &H1 'HWND_TOP=0
    End If
    
    ShowWindow lngHwnd, SW_RESTORE
End Sub


Public Sub UpdateSize()
'更新嵌入报告的窗口大小
    If mobjReportContainer Is Nothing Then Exit Sub

    If mblnIsPacsReport Then
        
        mfrmPacsReport.Left = 0
        mfrmPacsReport.Top = 0
        mfrmPacsReport.Width = mobjReportContainer.Width
        mfrmPacsReport.Height = mobjReportContainer.Height

'        '必须在这里设置一次窗口样式，否则可能造成嵌入的报告窗口包含标题栏
'        Call SetWindowStyle(mfrmPacsReport.hWnd)

    Else
        mobjReport.zlGetForm().Left = 0
        mobjReport.zlGetForm().Top = 0
        mobjReport.zlGetForm().Width = mobjReportContainer.Width
        mobjReport.zlGetForm().Height = mobjReportContainer.Height
    End If
End Sub


Public Sub OpenPacsReport()
On Error GoTo errHandle
    Dim i As Integer
    Dim j As Integer
    
    '如果没有选择行数据，则直接退出执行
    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, mobjOwner.Caption
        Exit Sub
    End If
    
    If Not mfrmPacsReportDock Is Nothing Then
        '先判断当前窗体是否是需要打开的窗体，如果不是，则查找窗口数组
        If mlngAdviceID = mfrmPacsReportDock.AdviceId And Not mfrmPacsReportDock.mblnClosed Then
            
            mfrmPacsReportDock.WindowState = 0  'normal
            mfrmPacsReportDock.ZOrder
            Exit Sub
        End If
    End If
    
    '查找窗口数组,找到需要打开的窗体，则通过Zorder把窗体显示到最前面
    If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
        For i = UBound(mobjPacsReportArry) To 1 Step -1
            If mobjPacsReportArry(i).mblnClosed Then
            
                If Not mfrmPacsReportDock Is Nothing Then
                    If mfrmPacsReportDock.AdviceId = mobjPacsReportArry(i).AdviceId Then
                        Set mfrmPacsReportDock = Nothing
                    End If
                End If
                
                '从数组中删除
                For j = i To UBound(mobjPacsReportArry) - 1
                    Set mobjPacsReportArry(j) = mobjPacsReportArry(j + 1)
                Next j
                
                ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) - 1) As frmReport
            ElseIf mlngAdviceID = mobjPacsReportArry(i).AdviceId Then
                Set mfrmPacsReportDock = mobjPacsReportArry(i)
                
                mfrmPacsReportDock.WindowState = 0  'normal
                mfrmPacsReportDock.ZOrder
                Exit Sub
            End If
        Next i
    End If
    
    
    '没有找到需要打开的窗体，且打开新窗体,并记录当前窗体
    Set mfrmPacsReportDock = New frmReport
    Set mfrmPacsReportDock.PacsCore = mobjPacsCore

    Call mfrmPacsReportDock.zlInitModule(mlngModule, mstrPrivs, mlngDeptID, mobjOwner, True)
    Call mfrmPacsReportDock.zlUpdateAdviceInf(mlngAdviceID, mlngSendNo, mlngStudyState, mblnMoved)
    
    Call mfrmPacsReportDock.zlEditReport
    
    If SafeArrayGetDim(mobjPacsReportArry) = 0 Then
        ReDim mobjPacsReportArry(1) As frmReport
    Else
        ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) + 1) As frmReport
    End If
    
    Set mobjPacsReportArry(UBound(mobjPacsReportArry)) = mfrmPacsReportDock
    Call OpenViewerWithReport
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Resume
End Sub


Private Sub OpenViewerWithReport()
'跟据参数打开报告后同时打开观片站，判断是否打开观片站
On Error GoTo errHandle
    
    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, mobjOwner.Caption
        Exit Sub
    End If
    
    If mblnShowImgAtReport And mstrStudyUID <> "" Then
        Dim intImageInverval As Integer
        
'        intImageInverval = Val(mfrmPACSImg.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)
        intImageInverval = 5        '图像加载间隔
        Call OpenViewer(1, mobjPacsCore, mlngAdviceID, False, mobjOwner, , , mblnLocalizerBackward, intImageInverval)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub BatPrint()
On Error GoTo errHandle
    Dim cbrControl As CommandBarControl
    Dim objReportPrint As New zlRichEPR.cDockReport
    Dim objPacsReport As New frmReport
    
    Dim strReturn As String
    Dim i As Integer
    Dim strReportString As String
    

    '选病人
    strReturn = frmDocPrintPatiList.Showfrm(mobjStudyList.DataGrid, mobjOwner, mblnCanPrint, mblnIsPacsReport, mlngDeptID)
    
    '循环调用报告打印
    '对于使用PACS报告编辑器打印的，调用PACS报告编辑器窗口来打印
    '返回值由"医嘱ID-是否PACS报告编辑器-执行科室ID|医嘱ID-是否PACS报告编辑器-执行科室ID|..."组成
    For i = 0 To UBound(Split(strReturn, "|"))
        strReportString = Split(strReturn, "|")(i)
        
        '判断是否使用PACS报告编辑器
        If Split(strReportString, "-")(1) = 1 Then  '使用PACS报告编辑器
            Call objPacsReport.zlInitModule(mlngModule, mstrPrivs, CLng(Split(strReportString, "-")(2)), mobjOwner)
            Call objPacsReport.zlUpdateAdviceInf(CLng(Split(strReportString, "-")(0)), 0, 0, False)
            Call objPacsReport.zlRefreshFace
            
            Call objPacsReport.zlMenu.zlCreateMenu(objPacsReport.cbrMain)
            Call objPacsReport.zlMenu.zlCreateToolBar(objPacsReport.cbrMain.Item(2))
            
            Call objPacsReport.zlMenu.zlExecuteMenu(conMenu_File_BatPrint)
            '需要AfterPrint吗？
        Else    '使用病历编辑器
            If objReportPrint.zlRefresh(CLng(Split(strReportString, "-")(0)), CLng(Split(strReportString, "-")(2)), , , True, mlngModule) > 0 Then
                Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_File_BatPrint + M_LNG_MENU_BASE_ID, , True)
                cbrControl.ID = cbrControl.ID - M_LNG_MENU_BASE_ID
                
                Call objReportPrint.zlExecuteCommandBars(cbrControl)
                cbrControl.ID = cbrControl.ID + M_LNG_MENU_BASE_ID
                
                RaiseEvent AfterPrinted(CLng(Split(strReportString, "-")(0)))
            End If
        End If
    Next
    
    Unload objPacsReport
    Unload objReportPrint.zlGetForm
    
    Set objPacsReport = Nothing
    Set objReportPrint = Nothing
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_ReportQuality(ByVal lngID As Long, ByVal strReportLevel As String, ByVal objStudyList As Object)
    On Error GoTo errHandle
    Dim strSql As String
    Dim strResult As String
    Dim strGrades() As String

    If mlngAdviceID <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, frmPacsMain.Caption
        Exit Sub
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

    strSql = "Zl_报告质量_Update(" & mlngAdviceID & ",'" & strResult & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "报告质量")

    strResult = strResult - 1
    strGrades = Split(strReportLevel, ",")
    
    If strResult <= UBound(strGrades) And strResult >= LBound(strGrades) Then
        objStudyList.CurText("报告质量") = strGrades(strResult)
        Call objStudyList.UpdateSourceData(mlngAdviceID, "报告质量", strResult + 1)
    End If
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub UpdateVideoCaptureState(ByVal lngWorkType As TWorkEventType, ByVal lngAdviceID As Long, ByVal other As Variant)
On Error GoTo errHandle
    Dim i As Long
    
    If mblnIsPacsReport Then
        '刷新嵌入报告中的缩略图图像或者视频采集的图像
        If Not mfrmPacsReport Is Nothing Then
            Call mfrmPacsReport.UpdateImageVideoState(lngWorkType, lngAdviceID, other)
        End If
    
    '    '刷新弹出报告窗口中的图像
    '    If Not mfrmPacsReportDock Is Nothing Then
    '        Call mfrmPacsReportDock.UpdateImageVideoState(lngWorkType, lngAdviceID, other)
    '    End If
    
        If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
            For i = LBound(mobjPacsReportArry) To UBound(mobjPacsReportArry)
                If Not mobjPacsReportArry(i) Is Nothing Then
                    Call mobjPacsReportArry(i).UpdateImageVideoState(lngWorkType, lngAdviceID, other)
                End If
            Next i
        End If
    Else
        If Not mobjReport Is Nothing Then mobjReport.RefPacsPic
    End If
errHandle:
End Sub


Public Sub ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'预览电子病历报告
On Error GoTo errHandle
    Dim cbrControl As CommandBarControl
    Dim lng医嘱ID As Long
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "select 医嘱ID FROM 病人医嘱报告 where 病历ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医嘱ID", CLng(报告ID))
    If rsTemp.EOF Then Exit Sub
    
    lng医嘱ID = Nvl(rsTemp!医嘱ID, 0)
    mobjReport.zlRefresh lng医嘱ID, mlngDeptID, False, , , mlngModule '以不可Edit方式刷新对像
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Help_Help, , True)
    
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '调用查阅报告
    cbrControl.ID = conMenu_Help_Help
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub RefreshReportImage()
    If mblnIsPacsReport = True Then
        Call mfrmPacsReport.RefPacsPic '刷新图片
        
        If Not mfrmPacsReportDock Is Nothing Then
            Call mfrmPacsReportDock.RefPacsPic '刷新图片
        End If
    Else
        Call mobjReport.RefPacsPic '刷新图片
    End If
End Sub






'AfterSaved
Private Sub mfrmPacsReport_AfterSaved(ByVal lngOrderID As Long, frmOwnerForm As Form, ByVal lngSaveType As Long)
    RaiseEvent AfterSaved(lngOrderID, frmOwnerForm, lngSaveType)
End Sub


Private Sub mfrmPacsReportDock_AfterSaved(ByVal lngOrderID As Long, frmOwnerForm As Form, ByVal lngSaveType As Long)
    RaiseEvent AfterSaved(lngOrderID, frmOwnerForm, lngSaveType)
End Sub

Private Sub mobjReport_AfterSaved(ByVal lngOrderID As Long, ByVal lngSaveType As Long)
    RaiseEvent AfterSaved(lngOrderID, mobjReport.zlGetForm, lngSaveType)
End Sub


'AfterClosed
Private Sub mfrmPacsReport_AfterClosed(ByVal lngOrderID As Long)
    RaiseEvent AfterClosed(lngOrderID)
End Sub

Private Sub mfrmPacsReportDock_AfterClosed(ByVal lngOrderID As Long)
    Call UpdateReporter(lngOrderID, "")
    
'    RaiseEvent AfterClosed(lngOrderID)
End Sub

Private Sub mobjReport_AfterClosed(ByVal lngOrderID As Long)
    Call UpdateReporter(lngOrderID, "")
    
'    RaiseEvent AfterClosed(lngOrderID)
End Sub

'BeforeEdit
Private Sub mfrmPacsReport_BeforeEdit(ByVal lngOrderID As Long)
    If CheckConcurrentReport(mobjOwner, lngOrderID) Then '检查是否有人正在操作报告
        Call UpdateReporter(lngOrderID, UserInfo.姓名)
    Else
        Call mfrmPacsReport.PromptModify(True)
    End If
    
    RaiseEvent BeforeEdit(lngOrderID)
End Sub

Private Sub mfrmPacsReportDock_BeforeEdit(ByVal lngOrderID As Long)
    If CheckConcurrentReport(mobjOwner, lngOrderID) Then '检查是否有人正在操作报告
        Call UpdateReporter(lngOrderID, UserInfo.姓名)
    Else
        Call mfrmPacsReportDock.PromptModify(True)
    End If
    
'    RaiseEvent BeforeEdit(lngOrderID)
End Sub


'AfterDeleted
Private Sub mfrmPacsReport_AfterDeleted(ByVal lngOrderID As Long)
    RaiseEvent AfterDeleted(lngOrderID)
End Sub

Private Sub mfrmPacsReportDock_AfterDeleted(ByVal lngOrderID As Long)
    RaiseEvent AfterDeleted(lngOrderID)
End Sub

Private Sub mobjReport_AfterDeleted(ByVal lngOrderID As Long)
    RaiseEvent AfterDeleted(lngOrderID)
End Sub

'AfterOpen
Private Sub mfrmPacsReport_AfterOpen()
    Exit Sub
End Sub


Private Sub mfrmPacsReportDock_AfterOpen()
'    Call OpenViewerWithReport
End Sub

Private Sub mobjReport_AfterOpen(ByVal intEditType As zlRichEPR.EditTypeEnum)
'    Call OpenViewerWithReport
End Sub


'AfterPrinted
Private Sub mfrmPacsReport_AfterPrinted(ByVal lngOrderID As Long)
    RaiseEvent AfterPrinted(lngOrderID)
End Sub

Private Sub mfrmPacsReportDock_AfterPrinted(ByVal lngOrderID As Long)
    RaiseEvent AfterPrinted(lngOrderID)
End Sub

Private Sub mobjReport_AfterPrinted(ByVal lngOrderID As Long)
    RaiseEvent AfterPrinted(lngOrderID)
End Sub





