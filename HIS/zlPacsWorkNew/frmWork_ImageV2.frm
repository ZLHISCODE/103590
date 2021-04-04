VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWork_ImageV2 
   BorderStyle     =   0  'None
   Caption         =   "影像观片"
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picView 
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   4560
      Width           =   4815
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   2055
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2415
         _Version        =   262147
         _ExtentX        =   4260
         _ExtentY        =   3625
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin MSComctlLib.ListView lvwSeq 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwImage 
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "frmWork_ImageV2.frx":0000
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmWork_ImageV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenuV2


Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请选择需要执行的检查记录。"
Private Const M_STR_MODULE_MENU_TAG As String = "影像"

Private mlngModule As Long
Private mstrPrivs As String
Private mlngDepartId As Long
Private mObjOwner As Object

Private mobjStudyInfo As clsStudyInfo
Private mObjNotify As IEventNotify


Private mlngCurImageCount As Long

'Private mintImageLocation As Integer            '记录图像所在的位置，0在中联数据库；1在新网数据库；2在新网数据库图像上传到云平台
Private mblnAutoOpenViewer As Boolean           '是否自动打开观片程序，ADViewer

Private mstrImageLevel As String                '影像质量等级串
Private mintImageLevel As Integer               '影像质量判定
Private mintViewHistoryImageDays As Integer     '自动打开历史图像天数

Private mblnShowPic As Boolean
Private mblnAddImage As Boolean                 '是否追加图像

Private iCurImageIndex As Integer
Private mintSelectAllSeq As Integer                 '0--无状态；1--选择全部序列；2--不选择全部序列
Private mintSelectAllImg As Integer                 '0--无状态；1--选择全部图像；2--不选择全部图像

Private mblnObserve As Boolean    '是否有观片基本权限   true是  false否

Private mblnUse3D As Boolean
Private mstr3DExeDir As String
Private mstr3DPara As String
Private mstr3DFunctions As String
Private mbln3DAutoDecompress As Boolean
Private mObjActiveMenuBar As CommandBars

Private mbyrFontState As Byte '字体状态，用于判断是否调整控件位置

Private mobjPacsCore As zl9PacsCore.clsViewer

Private mblnIsRefreshStudy As Boolean
Private mblnIsHistoryMode As Boolean

'科室ID
Property Get DeptId() As Long
    DeptId = mlngDepartId
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

'获取需要使用的外部对象
Property Get PacsCore() As Object
    Set PacsCore = mobjPacsCore
End Property

Property Set PacsCore(value As Object)
    Set mobjPacsCore = value
End Property


'获取菜单接口对象
Property Get zlMenu() As IWorkMenuV2
    Set zlMenu = Me
End Property



'接口实现部分**************************************************************************************************


Public Function IWorkMenuV2_zlBaseMenuID() As Long
End Function

Public Function IWorkMenuV2_zlExecuteCmd(ByVal lngCmdType As Long)
'执行菜单命令

End Function

Public Function IWorkMenuV2_zlGetModuleMenuId() As Long
'获取影像菜单的菜单ID
    IWorkMenuV2_zlGetModuleMenuId = conMenu_Img_Group
End Function



Public Function IWorkMenuV2_zlIsModuleMenu(ByVal strModuleName As String, objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单
    IWorkMenuV2_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenuV2_zlCreateMenu(ByVal strModuleName As String, objMenuBar As Object)
'创建影像记录对应的菜单
    
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim objGC As CommandBarControl
    
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    Set mObjActiveMenuBar = objMenuBar
    
    '删除影像质量子菜单
    Set cbrMenuBar = objMenuBar.FindControl(, conMenu_ManagePopup)
    Set cbrControl = cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_ImageQuality, , True)
    If Not cbrControl Is Nothing Then
        Call cbrControl.Delete
    End If

    Set cbrMenuBar = objMenuBar.FindControl(, conMenu_ManagePopup)
    With cbrMenuBar.CommandBar
        '创建影像质量菜单
        If CheckPopedom(mstrPrivs, "影像质控") Then
            Set objGC = cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_GChannel, , True)
            
            If objGC Is Nothing Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_ImageQuality, "影像质量", "", 0, False, .Controls.Count - 1)
            Else
                Set cbrControl = CreateModuleMenu(objGC.Parent.Controls, xtpControlPopup, conMenu_Manage_ImageQuality, "影像质量", "", 0, False, objGC.Index - 1)
            End If
            
            Call CreateSubordinateMenuTools(mstrImageLevel, cbrControl)
        End If
    End With
    
    If Not HasMenu(objMenuBar, conMenu_Img_Group) Then
        Set cbrMenuBar = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Img_Group, "影像", 3, False)
        cbrMenuBar.ID = conMenu_Img_Group
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Img_Look, "影像观片", "", 8111, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Img_Contrast, "影像对比", "", 8112, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Img_Look3D, "3D观片", "", 8115, False)
            
            '如果启用三维重建功能，则创建对应菜单
            If mblnUse3D = True Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Img_3D, "三维重建")  '.Add(xtpControlPopup, conMenu_Img_3D, "三维重建"): cbrControl.ID = conMenu_Img_3D
                    If mstr3DFunctions <> "" Then
                        str3DFuncs = Split(mstr3DFunctions, ",")
                        For i = 1 To UBound(str3DFuncs)
                            lng3DFunc = Val(str3DFuncs(i))
                            If lng3DFunc >= 1 And lng3DFunc <= 6 Then
                                Select Case lng3DFunc
                                    Case 1
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_VA, "容积重建")
                                    Case 2
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                    Case 3
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                    Case 4
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_VE, "虚拟内窥镜")
                                    Case 5
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_SA, "表面重建")
                                    Case 6
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_PF, "灌注成像")
                                End Select
                            End If
                        Next i
                    End If
            End If
             
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Img_Delete, "影像全删", "", 8113, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Img_Query, "影像获取(Q/R)", "", 8111, False)
            
            If gblnUseXinWangView = True Then
                '判断是否使用新版观片
                'Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_FilmPrevew, "胶片预览", "", 0, True)
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_FilmPrint, "胶片打印", "", 0, False)
                'Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_FilmDelete, "胶片删除", "", 0, False)
            End If
            
            If CheckPopedom(mstrPrivs, "技师执行") Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_TechDoctorExecute, "技师执行", "指定当前检测的检查技师", 807, True)
            End If
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ChangeDevice, "更换影像类型", "", 3203, False)
            
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Show, "影像显示", "显示当前序列影像缩略图", 3061, True): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Expend_AllCollapse, "全选序列", "选中当前所有序列", 3010, False): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Expend_AllExpend, "全清序列", "清除选中当前所有序列", 3004, False): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SelectAllImages, "全选图像", "选中当前所有图像", 227, False): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_UnSelectAllImages, "全清图像", "清除选中当前所有图像", 229, False): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ReverseSelectImages, "反选图像", "反向选择所有图像", 3012, False): cbrControl.Style = xtpButtonIconAndCaption
                
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_DeleteSeries, "删除选定序列", "删除选择的序列", 0, True): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_DeleteImage, "删除选定图像", "删除选择的图像", 0, False): cbrControl.Style = xtpButtonIconAndCaption
                
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_DevSet, "影像设备设置", "刷新当前病人图像序列", 181, True): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_RefreshImg, "刷新图像", "刷新当前病人图像序列", 791, True): cbrControl.Style = xtpButtonIconAndCaption
        End With
    End If
End Sub

Public Sub IWorkMenuV2_zlCreateToolBar(ByVal strModuleName As String, objToolBar As Object)
'创建工具栏
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim lngIndex As Long
        
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    '删除影像质量工具栏
    Set cbrControl = objToolBar.FindControl(, conMenu_Manage_ImageQuality)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete

    '创建影像质量工具栏
    If CheckPopedom(mstrPrivs, "影像质控") Then
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_ImageQuality, "影像质量", "影像质量", 3061, False, mObjActiveMenuBar.FindControl(, conMenu_Manage_Result).Index + 1)
        Call CreateSubordinateMenuTools(mstrImageLevel, cbrControl)
    End If
    
    If HasMenu(objToolBar, conMenu_Img_Look) Then Exit Sub
    
    Set cbrLogOut = objToolBar.FindControl(, conMenu_img_ContrastView)
    
    lngIndex = 2
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index
'
'    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Img_Look, "观片", "影像观片", 8111, True, lngIndex + 1)
'    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Img_Contrast, "对比", "观片对比", 8112, False, lngIndex + 2)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Img_Look3D, "3D观片", "3D观片", 8115, False, lngIndex + 1)
    
    If CheckPopedom(mstrPrivs, "技师执行") Then
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Manage_TechDoctorExecute, "技师执行", "指定当前检测的检查技师", 807, False, lngIndex + 2)
    End If
    
    '如果启用三维重建功能，则创建对应菜单
    If mblnUse3D = True Then
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButtonPopup, conMenu_Img_3D, "三维", "三维重建", 8115, False, lngIndex + 3)
            If mstr3DFunctions <> "" Then
                str3DFuncs = Split(mstr3DFunctions, ",")
                For i = 1 To UBound(str3DFuncs)
                    lng3DFunc = Val(str3DFuncs(i))
                    If lng3DFunc >= 1 And lng3DFunc <= 6 Then
                        Select Case lng3DFunc
                            Case 1
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_VA, "容积重建")
                            Case 2
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                            Case 3
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                            Case 4
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_VE, "虚拟内窥镜")
                            Case 5
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_SA, "表面重建")
                            Case 6
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_PF, "灌注成像")
                        End Select
                    End If
                Next i
            End If
    End If
    
    If gblnUseXinWangView = True Then
        '判断是否使用新版观片
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Manage_FilmPrint, "胶片打印", "", 3202, False, lngIndex + 4)
    End If
End Sub

Private Sub CreateSubordinateMenuTools(ByVal strImageLevel As String, ByVal cbrControl As CommandBarControl)
'创建下级菜单和工具栏
    Dim cbrPopControl As CommandBarControl
    Dim intTxtLen As Integer
    Dim i As Integer
    
    intTxtLen = Len(strImageLevel) - Len(Replace(strImageLevel, ",", "")) + 1
    For i = 1 To 4
        If i <= intTxtLen Then
            If Trim(Split(strImageLevel, ",")(i - 1)) <> "" Then
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, _
                    Decode(i, 1, conMenu_Manage_ImageFirst, 2, conMenu_Manage_ImageSecond, 3, conMenu_Manage_ImageThird, 4, conMenu_Manage_ImageFourth), Trim(Split(strImageLevel, ",")(i - 1)), "", 0, False)
            End If
        End If
    Next i
End Sub

Public Sub IWorkMenuV2_zlClearMenu(ByVal strModuleName As String)
'清除所创建的菜单
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    
    If mObjActiveMenuBar Is Nothing Then Exit Sub
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_Group)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    '删除影像质量子菜单
    Set cbrMenuBar = mObjActiveMenuBar.FindControl(, conMenu_ManagePopup)
    Set cbrControl = cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_ImageQuality)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub


Public Sub IWorkMenuV2_zlClearToolBar(ByVal strModuleName As String)
'清除创建的工具栏
    Dim cbrControl As CommandBarControl
    
    If mObjActiveMenuBar Is Nothing Then Exit Sub
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_Look)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_Contrast)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_Look3D)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Manage_TechDoctorExecute)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_3D)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    If gblnUseXinWangView = True Then
    '判断是否使用新版观片
        Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Manage_FilmPrint)
        If Not cbrControl Is Nothing Then Call cbrControl.Delete
    End If
    
    '删除影像质量工具栏
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Manage_ImageQuality)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub

Public Sub IWorkMenuV2_zlExecuteMenu(ByVal strModuleName As String, ByVal lngMenuId As Long)
'根据菜单ID执行对应功能
    
    Select Case lngMenuId
        Case conMenu_Img_Look                           '观片
            Call Menu_Img_观片
        Case conMenu_Img_Contrast                       '对比观片
            Call Menu_Img_对比观片
        Case conMenu_Img_Look3D                         '3D观片
            Call Menu_Img_3D观片
        Case conMenu_Manage_FilmPrint                   '胶片打印
            Call Menu_Manage_FilmPrint
        Case conMenu_Img_3D_MMPR                        '三维重建，MMPR
            Call sub三维重建("MMPR")
        Case conMenu_Img_3D_MPR                         '三维重建，MPR
            Call sub三维重建("MPR")
        Case conMenu_Img_3D_PF                          '三维重建,灌注成像
            Call sub三维重建("PF")
        Case conMenu_Img_3D_SA                          '三维重建，表面重建
            Call sub三维重建("SA")
        Case conMenu_Img_3D_VA                          '三维重建，容积重建
            Call sub三维重建("VA")
        Case conMenu_Img_3D_VE                          '三维重建，虚拟内窥镜
            Call sub三维重建("VE")
        Case conMenu_Img_Delete                         '图象删除
            Call Menu_Img_图象删除
        Case conMenu_Img_Query                          '从设备获取图象
            Call Menu_Img_获取图像
        Case conMenu_Manage_TechDoctorExecute           '技师执行
            Call Menu_Img_技师执行
        Case conMenu_Manage_ChangeDevice                '更换影像类型
            Call Menu_Img_更换检查设备
        Case conMenu_View_Show          '显示图像
            mblnShowPic = Not mblnShowPic
            Call zlMenuClick("影像显示")
        Case conMenu_View_Expend_AllCollapse    '全选序列
            Call zlMenuClick("全选序列")
        Case conMenu_View_Expend_AllExpend      '全清序列
            Call zlMenuClick("全清序列")
        Case conMenu_Manage_SelectAllImages     '全选图像
            Call zlMenuClick("全选图像")
        Case conMenu_Manage_UnSelectAllImages   '全清图像
            Call zlMenuClick("全清图像")
        Case conMenu_Manage_ReverseSelectImages '反选图像
            Call zlMenuClick("反选图像")
        Case conMenu_Manage_DeleteSeries        '删除序列
            Call zlMenuDeleteImageClick(lngMenuId)
        Case conMenu_Manage_DeleteImage         '删除图像
            Call zlMenuDeleteImageClick(lngMenuId)
        Case conMenu_Cap_DevSet                 '影像设备设置
            frmPACSImageDeviceSetup.Show vbModal, mObjOwner
        Case conMenu_Manage_RefreshImg          '刷新图像
            Call zlRefreshFace(mobjStudyInfo, True)
        Case conMenu_Manage_ImageFirst, conMenu_Manage_ImageSecond, conMenu_Manage_ImageThird, conMenu_Manage_ImageFourth
            Call Menu_Manage_影像质量(lngMenuId, mstrImageLevel)
    End Select
End Sub

Public Sub IWorkMenuV2_zlUpdateMenu(ByVal strModuleName As String, Control As XtremeCommandBars.ICommandBarControl)
'更新菜单
    If mobjStudyInfo Is Nothing Then
        Control.Enabled = False
        Exit Sub
    End If
    
    Select Case Control.ID
        Case conMenu_Img_Look       '观片，当前检查有图像，或者是患者有历史检查，则可以观片
            Control.Enabled = (mobjStudyInfo.strStudyUID <> "") Or mintViewHistoryImageDays > 0
            
            
        Case conMenu_Img_Contrast   '观片对比，只有中联PACS才显示
            Control.Enabled = (mobjStudyInfo.strStudyUID <> "")
            Control.Visible = IIf(mobjStudyInfo.intImageLocation = 0, True, False)
        
        Case conMenu_Img_Look3D     '3D观片，只有新网PACS才显示
            Control.Enabled = mobjStudyInfo.strStudyUID <> "" And mlngCurImageCount >= 50
            Control.Visible = (mobjStudyInfo.intImageLocation <> 0)
            
        Case conMenu_Manage_FilmPrint                    '胶片打印
            Control.Visible = CheckPopedom(mstrPrivs, "胶片按需打印")
            
        Case conMenu_Img_3D         '三维重建
            If CheckPopedom(mstrPrivs, "三维重建操作") And mblnUse3D = True Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
            
            If Control.Visible = True Then Control.Enabled = mobjStudyInfo.strStudyUID <> ""
            
        Case conMenu_Img_Delete '清除图像，图像在云平台，不显示按钮
            If Not CheckPopedom(mstrPrivs, "清除图像") Or (mobjStudyInfo.intImageLocation = 2) Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
            
            If Control.Visible = True Then Control.Enabled = mobjStudyInfo.strStudyUID <> ""
            
            Control.Enabled = Not mblnIsHistoryMode
            
        Case conMenu_Img_Query ',获取图像，只有中联PACS才显示
            If (Not CheckPopedom(mstrPrivs, "清除图像")) Or (mobjStudyInfo.intImageLocation <> 0) Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
            
            If Control.Visible Then Control.Enabled = mobjStudyInfo.intStep > 1
            
        Case conMenu_Manage_ChangeDevice    '更改影像设备类型
                If mobjStudyInfo.strImgType = "CR" Or _
                    mobjStudyInfo.strImgType = "DR" Or _
                    mobjStudyInfo.strImgType = "DX" Or _
                    mobjStudyInfo.strImgType = "RF" Then
                    Control.Enabled = True
                Else
                    Control.Enabled = False
                End If
                
                Control.Enabled = Control.Enabled And Not mblnIsHistoryMode
                
        Case conMenu_Manage_TechDoctorExecute   '技师执行
            If mobjStudyInfo.blnIsTechincalSure Then Control.Caption = "技师取消" Else Control.Caption = "技师执行"
            
            If mobjStudyInfo.intStep >= 2 And mobjStudyInfo.intStep < 5 Then
                Control.Enabled = True
                
                If mobjStudyInfo.blnIsTechincalSure Then
                    Control.Enabled = UserInfo.姓名 = mobjStudyInfo.strDoDoctor Or CheckPopedom(mstrPrivs, "取消技师执行")
                End If
            Else
                Control.Enabled = False
            End If
            
            Control.Enabled = Control.Enabled And Not mblnIsHistoryMode
            
        Case conMenu_Manage_DeleteSeries    '删除选定序列
            Control.Enabled = lvwSeq.ListItems.Count > 0 And Me.Visible And Not mblnIsHistoryMode
        Case conMenu_Manage_DeleteImage     '删除选定图像
            Control.Enabled = lvwImage.ListItems.Count > 0 And Me.Visible And Not mblnIsHistoryMode
        Case conMenu_View_Show, conMenu_View_Expend_AllCollapse, conMenu_View_Expend_AllExpend  '图像显示，全选序列，全清序列
            Control.Enabled = lvwSeq.ListItems.Count > 0 And Me.Visible
            Control.Visible = (mobjStudyInfo.intImageLocation = 0) And Me.Visible
            Control.Checked = Me.cbrMain.FindControl(, Control.ID).Checked
        Case conMenu_Manage_SelectAllImages, conMenu_Manage_UnSelectAllImages, conMenu_Manage_ReverseSelectImages   '全选图像，全清图像，反选图像
            Control.Enabled = lvwImage.ListItems.Count > 0 And Me.Visible
            Control.Visible = (mobjStudyInfo.intImageLocation = 0) And Me.Visible
            Control.Checked = Me.cbrMain.FindControl(, Control.ID).Checked
        Case conMenu_Img_Group, conMenu_Img_Query, conMenu_View_Refresh, conMenu_Cap_DevSet '影像，影像获取，图像刷新
            Control.Enabled = True
        Case conMenu_Manage_ImageFirst, conMenu_Manage_ImageSecond, conMenu_Manage_ImageThird, conMenu_Manage_ImageFourth, conMenu_Manage_ImageQuality
            If Not CheckPopedom(mstrPrivs, "影像质控") Or mintImageLevel = 0 Then
                Control.Visible = False
            ElseIf (mobjStudyInfo.intStep >= 3 And mobjStudyInfo.intStep <= 5) Or mobjStudyInfo.intStep = -1 Then
                Control.Visible = True
                Control.Enabled = mobjStudyInfo.strStudyUID <> ""
            Else
                Control.Visible = True
                Control.Enabled = False
            End If
    End Select
End Sub

Public Sub IWorkMenuV2_zlPopupMenu(ByVal strModuleName As String, objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单
    Dim objControl As CommandBarControl
    Dim objMenuControl As CommandBarControl
    
    For Each objMenuControl In mObjActiveMenuBar.ActiveMenuBar.Controls
        If objMenuControl.ID = conMenu_Img_Group And objMenuControl.type = xtpControlPopup Then
            For Each objControl In objMenuControl.CommandBar.Controls
                If objControl.ID = conMenu_Img_Look Or _
                   objControl.ID = conMenu_Img_Contrast Or _
                   objControl.ID = conMenu_Img_Look3D Or _
                   objControl.ID = conMenu_Img_Delete Or _
                   objControl.ID = conMenu_Img_Query Or _
                   objControl.ID = conMenu_Manage_TechDoctorExecute Or _
                   objControl.ID = conMenu_Manage_ChangeDevice Or _
                   objControl.ID = conMenu_Manage_FilmPrint Then
                   
                    objControl.Copy objPopup
                    
                End If
            Next
        End If
    Next
End Sub

Public Sub IWorkMenuV2_zlRefreshSubMenu(ByVal strModuleName As String, objMenuBar As Object)
'刷新弹出的子菜单
    Exit Sub
End Sub


'**********************************************************************************************************************

Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
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


Public Sub zlInitModule(objNotify As IEventNotify, ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long)
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngDepartId = lngDepartId
    
    Set mObjNotify = objNotify
    
    If Not objNotify Is Nothing Then Set mObjOwner = objNotify.Owner
    
    mblnUse3D = Val(zlDatabase.GetPara("启用三维重建", glngSys, lngModule, 0))
    mstr3DExeDir = zlDatabase.GetPara("3D程序路径", glngSys, lngModule, "")
    mstr3DPara = zlDatabase.GetPara("3D参数", glngSys, lngModule, "")
    mstr3DFunctions = zlDatabase.GetPara("3D功能", glngSys, lngModule, "")
    mbln3DAutoDecompress = Val(zlDatabase.GetPara("3D自动解压缩", glngSys, lngModule, 0))
    mstrImageLevel = nvl(GetDeptPara(mlngDepartId, "影像质量等级", "甲,乙"))
    mintImageLevel = Val(GetDeptPara(mlngDepartId, "影像质量判定", "0"))
    mintViewHistoryImageDays = Val(GetDeptPara(mlngDepartId, "自动打开历史图像天数", 0))
    If mintViewHistoryImageDays > 15 Or mintViewHistoryImageDays <= 0 Then
        mintViewHistoryImageDays = 1
    End If
End Sub



Public Function zlRefreshFace(objStudyInfo As clsStudyInfo, _
    Optional blnForceRefresh As Boolean = False, Optional ByVal blnIsHistory As Boolean = False) As Boolean
On Error GoTo DBError

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    mblnIsHistoryMode = blnIsHistory
       
    If Not mobjStudyInfo Is Nothing And Not objStudyInfo Is Nothing Then
        If mblnIsRefreshStudy = True Then    '已经刷新过则需要判断检查信息相同以及是否进行强制刷新
            If mobjStudyInfo.IsEquals(objStudyInfo) And blnForceRefresh = False Then Exit Function
        Else
            blnForceRefresh = True
        End If
    End If
        
    
    mblnShowPic = False

    Set mobjStudyInfo = objStudyInfo

    '转出的影像不能保存报告
    If mobjStudyInfo.blnMoved Then
        mstrPrivs = Replace(mstrPrivs, "图像操作处理", "")
        mstrPrivs = Replace(mstrPrivs, "图像标注测量", "")
        mstrPrivs = Replace(mstrPrivs, "清除图像", "")
    End If
    
    Call ShowSeqImg
    
    mblnIsRefreshStudy = True
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Menu_Manage_影像质量(ByVal lngID As Long, ByVal strImageLevel As String)
On Error GoTo errhandle
    Dim strSQL As String
    Dim strResult As String
    Dim strGrades() As String

    If mobjStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, frmPacsMain.Caption
        Exit Sub
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_QUALITYTAG, 0, mobjStudyInfo.lngAdviceId)
    
    Select Case lngID
        Case conMenu_Manage_ImageFirst
            strResult = 1
        Case conMenu_Manage_ImageSecond
            strResult = 2
        Case conMenu_Manage_ImageThird
            strResult = 3
        Case conMenu_Manage_ImageFourth
            strResult = 4
    End Select

    strSQL = "Zl_影像质量_Update(" & mobjStudyInfo.lngAdviceId & ",'" & strResult & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "影像质量")
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_QUALITYTAG, 1, mobjStudyInfo.lngAdviceId, strResult)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Img_观片()
On Error GoTo errhandle
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD mObjOwner, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    '刷新界面
    Call zlRefreshFace(mobjStudyInfo, False)
    
    Call zlMenuClick("影像处理")
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Img_3D观片()
On Error GoTo errhandle
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD mObjOwner, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    '刷新界面
    Call zlRefreshFace(mobjStudyInfo, False)
    
    Call zlMenuClick("影像3D观片")
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Img_对比观片()
On Error GoTo errhandle
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD mObjOwner, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    '刷新界面
    Call zlRefreshFace(mobjStudyInfo, False)
    
    Call zlMenuClick("影像对比")
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_FilmPrint()
'胶片打印
On Error GoTo errhandle
    Dim blnPrintResult As Boolean
    
    '判断是否具有相应操作权限
    If Not CheckPopedom(mstrPrivs, "胶片按需打印") Then
        MsgBoxD Me, "您不具备胶片打印权限，请联系管理员。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_XWFILMPRINT, 0, mobjStudyInfo.lngAdviceId)
    
    blnPrintResult = XWShowFilmPrintWind(mobjStudyInfo.lngAdviceId, Me)
    
    If blnPrintResult = True Then
        '发送胶片打印消息到主窗口中
        If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_XWFILMPRINT, 1, mobjStudyInfo.lngAdviceId)
    End If
    
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Img_图象删除()
On Error GoTo errhandle
    Dim rsTemp As ADODB.Recordset
    Dim blnIsCancel As Boolean
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    '如果图像在云平台mintImageLocation=2，不显示“图像删除”按钮，不允许删除图像
    If mobjStudyInfo.intImageLocation = 1 Then
        '图像在新网PACS，调用ArchiveManager删除图像
        Call subXWShowArchiveManager(1)
    ElseIf mobjStudyInfo.intImageLocation = 0 Then    '图像在中联PACS
        If Not mObjNotify Is Nothing Then
            Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 0, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, blnIsCancel)
            If blnIsCancel Then Exit Sub
        End If
        
        Call zlRefreshFace(mobjStudyInfo, False)
            
        gstrSQL = "select 检查UID from 影像检查记录 where 医嘱ID =[1] and  发送号 = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取检查UID", mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo)
        
        If rsTemp.EOF Then Exit Sub
            
        If MsgBoxD(Me, "是否确认要删除该检查的所有影像？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub

        
        '删除影像文件和目录
        RemoveCheckImages mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo
        
        gstrSQL = "ZL_影像检查_PhotoDelete(" & mobjStudyInfo.lngAdviceId & "," & mobjStudyInfo.lngSendNo & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
        Call ClearListData
        
        '如果最后一个序列也被删除了,应该刷新列表
        If lvwSeq.ListItems.Count = 0 Then
            If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, -1, hwnd)
        Else
            If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, , hwnd)
        End If
    End If
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'功能:重新设置工作站窗体的字体大小
    
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    
    Me.FontSize = bytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("ListView")
            objCtrl.Font.Size = bytFontSize
            objCtrl.Font.Name = "微软雅黑"
        Case UCase("TabStrip") '页面控件
            objCtrl.Font.Size = bytFontSize
        Case UCase("Label")
            objCtrl.FontSize = bytFontSize
            objCtrl.Height = TextHeight("罗") + 20
        Case UCase("vsFlexGrid")
            objCtrl.FontSize = bytFontSize
        Case UCase("ucFlexGrid")
            objCtrl.DataGrid.Cell(flexcpFontSize, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = bytFontSize
            objCtrl.DataGrid.FontSize = bytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = bytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = bytFontSize
            objCtrl.Width = TextWidth("罗冠" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.Width = TextWidth("罗冠" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = bytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("罗") * 1.5
        Case UCase("textBox")
          objCtrl.FontSize = bytFontSize
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.options.Font = CtlFont
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = bytFontSize
        End Select
    Next
    
    Call lvwSeq.Refresh
    
End Sub




Private Sub ClearListData()
'删除界面列表中的数据
    lvwSeq.ListItems.Clear
    lvwImage.ListItems.Clear
    DViewer.Images.Clear
End Sub


Private Sub Menu_Img_获取图像()
On Error GoTo errhandle
    Dim strImageDeviceNumber As String, rsTemp As ADODB.Recordset

    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_GETIMAGE, 0, mobjStudyInfo.lngAdviceId)
    
    Call zlRefreshFace(mobjStudyInfo, False)
    
    strImageDeviceNumber = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSImageDeviceSetup", "默认影像设备", "")
    
    '没有默认设备时处理
    If strImageDeviceNumber = "" Then
        If MsgBoxD(Me, "没有设置默认影像检查设备！是否现在设置？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            frmPACSImageDeviceSetup.Show vbModal, Me
            strImageDeviceNumber = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSImageDeviceSetup", "默认影像设备", "")
            If strImageDeviceNumber = "" Then Exit Sub
        End If
    End If
    
    gstrSQL = "select 设备号,设备名, IP地址,端口号,本地AE,设备AE from 影像设备目录 where 设备号 = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(Mid(strImageDeviceNumber, 2)))
    
    '当默认设备被删除后重新设置
    If rsTemp.EOF = True Then
        MsgBoxD Me, "默认设备已被删除，请重新设置！", vbInformation, gstrSysName
        frmPACSImageDeviceSetup.Show vbModal, Me
        Exit Sub
    End If
        
    '先判断设备的AE，端口是否被正确设置了,未设置好则提示并退出
    If IsNull(rsTemp("端口号")) Or IsNull(rsTemp("设备AE")) Or IsNull(rsTemp("本地AE")) Then
        MsgBoxD Me, "请到“影像设备目录”模块中，设置Q/R查询使用的设备端口号，设备AE和本地AE。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    frmPACSGetDeviceImage.ShowMe Me, rsTemp("IP地址"), rsTemp("端口号"), rsTemp("设备名"), rsTemp("本地AE"), rsTemp("设备AE"), mobjStudyInfo.lngAdviceId
        
        
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_GETIMAGE, 1, mobjStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Img_技师执行()
'技师执行或技师取消
On Error GoTo errhandle
    Dim strSQL As String
    Dim intResult As Integer '0-取消；1-执行
        
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_TECHDO, 0, mobjStudyInfo.lngAdviceId)
    
    If mobjStudyInfo.blnIsTechincalSure Then     '技师取消
        strSQL = "Zl_影像技师执行('" & UserInfo.姓名 & "'," & mobjStudyInfo.lngAdviceId & ",1)"
        Call zlDatabase.ExecuteProcedure(strSQL, "技师取消")
        
        mobjStudyInfo.blnIsTechincalSure = False
        
        intResult = 0
    Else
        If mobjStudyInfo.strDoDoctor <> UserInfo.姓名 Then
            If Not MsgBoxD(Me, "当前人员与指定的检查技师不相同," & vbCrLf & "确定要继续执行吗？", vbYesNo, "技师执行") = vbNo Then
                strSQL = "Zl_影像技师执行('" & UserInfo.姓名 & "'," & mobjStudyInfo.lngAdviceId & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, "技师执行")
                
                mobjStudyInfo.blnIsTechincalSure = True
                mobjStudyInfo.strDoDoctor = UserInfo.姓名
                
                intResult = 1
            End If
        Else
            strSQL = "Zl_影像技师执行('" & UserInfo.姓名 & "'," & mobjStudyInfo.lngAdviceId & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "技师执行")
            
            mobjStudyInfo.blnIsTechincalSure = True
            mobjStudyInfo.strDoDoctor = UserInfo.姓名
            
            intResult = 1
        End If
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_TECHDO, 1, mobjStudyInfo.lngAdviceId, intResult)

    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Img_更换检查设备()
On Error GoTo errhandle
    Dim strModality As String
    Dim rResult As VbMsgBoxResult
    Dim strSQL As String
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_CHANGEDEVICE, 0, mobjStudyInfo.lngAdviceId)
    
    frmChangeDevice.ShowMe UCase(mobjStudyInfo.strImgType), Me
    strModality = frmChangeDevice.strDeviceType
    
    If strModality <> "" Then
        strSQL = "Zl_影像检查_影像类别(" & mobjStudyInfo.lngAdviceId & "," & mobjStudyInfo.lngSendNo & ",'" & strModality & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_CHANGEDEVICE, 1, mobjStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub sub三维重建(strCommand As String)
    Dim strImageDir As String
    
    Call zlRefreshFace(mobjStudyInfo, False)
    
    '组织三维重建需要的图像
    strImageDir = ZLfun3DImgProcess(mbln3DAutoDecompress)
    If strImageDir <> "" Then
        Call sub3DProcess(strCommand, strImageDir)
    End If
End Sub


Private Sub sub3DProcess(strCommand As String, strImageDir As String)
On Error GoTo errhandle
    Dim str3DCommand As String
    
    '组织三维重建语句
    str3DCommand = mstr3DExeDir & " " & mstr3DPara & " " & strCommand & " " & strImageDir
    
    Shell str3DCommand
    
errhandle:
End Sub


'执行菜单命令
Public Sub zlMenuClick(mnuClick As String)
    
    mblnAddImage = False
    Select Case mnuClick
        Case "影像处理"
            DViewer_DblClick
        Case "影像对比"
            mblnAddImage = True
            DViewer_DblClick
        Case "影像3D观片"
            Call Open3DViewer(mobjStudyInfo.lngAdviceId, Me, mobjStudyInfo.blnMoved)
        Case "影像显示"
            If Not lvwImage.SelectedItem Is Nothing Then ShowLvwImage lvwImage.SelectedItem
        Case "全选序列"
            If mintSelectAllSeq = 0 Or mintSelectAllSeq = 2 Then
                mintSelectAllSeq = 1
            ElseIf mintSelectAllSeq = 1 Then
                mintSelectAllSeq = 0
            End If
            Call subSetMenuState
            SelectAllSeq True
        Case "全清序列"
            If mintSelectAllSeq = 0 Or mintSelectAllSeq = 1 Then
                mintSelectAllSeq = 2
            ElseIf mintSelectAllSeq = 2 Then
                mintSelectAllSeq = 0
            End If
            Call subSetMenuState
            SelectAllSeq False
        Case "全选图像"
            If mintSelectAllImg = 0 Or mintSelectAllImg = 2 Then
                mintSelectAllImg = 1
            ElseIf mintSelectAllImg = 1 Then
                mintSelectAllImg = 0
            End If
            Call subSetMenuState
            SelectAllImg True
        Case "全清图像"
            If mintSelectAllImg = 0 Or mintSelectAllImg = 1 Then
                mintSelectAllImg = 2
            ElseIf mintSelectAllImg = 2 Then
                mintSelectAllImg = 0
            End If
            Call subSetMenuState
            SelectAllImg False
        Case "反选图像"
            Dim i As Integer
            With lvwImage
                For i = 1 To .ListItems.Count
                    .ListItems(i).Checked = Not .ListItems(i).Checked
                Next
            End With
            Call WriteSelectdImages(lvwImage.tag)
    End Select
End Sub

Private Sub subSetMenuState()
On Error Resume Next
    If mblnShowPic Then
        Me.cbrMain.FindControl(, conMenu_View_Show).Checked = True
    Else
        Me.cbrMain.FindControl(, conMenu_View_Show).Checked = False
    End If
    
    If mintSelectAllSeq = 0 Then            '0--无状态
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllCollapse).Checked = False
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllExpend).Checked = False
    ElseIf mintSelectAllSeq = 1 Then        '1--选择全部序列
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllCollapse).Checked = True
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllExpend).Checked = False
    ElseIf mintSelectAllSeq = 2 Then        '2--不选择全部序列
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllCollapse).Checked = False
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllExpend).Checked = True
    End If
    
    If mintSelectAllImg = 0 Then            '0--无状态
        Me.cbrMain.FindControl(, conMenu_Manage_SelectAllImages).Checked = False
        Me.cbrMain.FindControl(, conMenu_Manage_UnSelectAllImages).Checked = False
    ElseIf mintSelectAllImg = 1 Then        '1--选择全部图像
        Me.cbrMain.FindControl(, conMenu_Manage_SelectAllImages).Checked = True
        Me.cbrMain.FindControl(, conMenu_Manage_UnSelectAllImages).Checked = False
    ElseIf mintSelectAllImg = 2 Then        '2--不选择全部图像
        Me.cbrMain.FindControl(, conMenu_Manage_SelectAllImages).Checked = False
        Me.cbrMain.FindControl(, conMenu_Manage_UnSelectAllImages).Checked = True
    End If
    
    
End Sub

Private Sub SelectAllSeq(ByVal blnSelect As Boolean)
    Dim i As Integer
    With lvwSeq
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next

        '图像在中联PACS
        If mobjStudyInfo.intImageLocation = 0 Then
            If Not lvwSeq.SelectedItem Is Nothing Then
                ShowImageList lvwSeq.SelectedItem
            Else
                ShowImageList Nothing
            End If
        End If
    End With
End Sub

Private Sub SelectAllImg(ByVal blnSelect As Boolean)
    Dim i As Integer
    With lvwImage
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next
    End With
    Call WriteSelectdImages(lvwImage.tag)
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_Show          '显示图像
            mblnShowPic = Not mblnShowPic
            Control.Checked = mblnShowPic
            Call zlMenuClick("影像显示")
        Case conMenu_View_Expend_AllCollapse    '全选序列
            Call zlMenuClick("全选序列")
        Case conMenu_View_Expend_AllExpend      '全清序列
            Call zlMenuClick("全清序列")
        Case conMenu_Manage_SelectAllImages     '全选图像
            Call zlMenuClick("全选图像")
        Case conMenu_Manage_UnSelectAllImages   '全清图像
            Call zlMenuClick("全清图像")
        Case conMenu_Manage_ReverseSelectImages '反选图像
            Call zlMenuClick("反选图像")
        Case conMenu_View_Refresh
            Call zlRefreshFace(mobjStudyInfo, True)
        Case conMenu_Manage_DeleteSeries        '删除序列
            Call zlMenuDeleteImageClick(Control.ID)
        Case conMenu_Manage_DeleteImage         '删除图像
            Call zlMenuDeleteImageClick(Control.ID)
    End Select
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
Exit Sub
    Select Case Control.ID
        Case conMenu_View_Expend_AllCollapse, conMenu_View_Expend_AllExpend   '全选序列，全清序列，

            Control.Enabled = lvwSeq.ListItems.Count > 0
            Control.Checked = False
            
        Case conMenu_Manage_SelectAllImages, conMenu_Manage_UnSelectAllImages, conMenu_Manage_ReverseSelectImages '全选图像，全清图像，反选图像
            Control.Enabled = lvwSeq.ListItems.Count > 0
            Control.Visible = (mobjStudyInfo.intImageLocation = 0)
            Control.Checked = False
            
        Case conMenu_View_Show
            Control.Enabled = lvwSeq.ListItems.Count > 0
            Control.Visible = (mobjStudyInfo.intImageLocation = 0)
            Control.Checked = mblnShowPic
            
        Case conMenu_Manage_ImageInterval   '图像间隔
            Control.Visible = (mobjStudyInfo.intImageLocation = 0)
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = lvwSeq.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = lvwImage.hwnd
    ElseIf Item.ID = 3 Then
        Item.Handle = picView.hwnd
    End If
End Sub

Private Sub DViewer_DblClick()
'显示观片站
    Dim strSerials As String, strSeqUID As String
    Dim Item As MSComctlLib.ListItem
    Dim intImageInverval As Integer
    Dim strImages As String
    Dim rsTemp As ADODB.Recordset
    Dim strFtpURL As String
    
    On Error GoTo CallError
       
    '图像在新网数据库管理
    If mobjStudyInfo.intImageLocation = 1 Or mobjStudyInfo.intImageLocation = 2 Then
        strSerials = ""
        
        If mobjStudyInfo.intImageLocation = 1 Then
            If lvwSeq.SelectedItem Is Nothing Then Exit Sub '当前检查没有图像，就退出
            
            
            For Each Item In lvwSeq.ListItems
                strSeqUID = Mid(Item.Key, 2)
                If Item.Checked Then
                    '只有当前序列被勾选了，而且选择可部分图象或者全部图象，才打开该序列
                    strSerials = strSerials & ",'" & strSeqUID & "'"
                End If
            Next
            
            strSerials = Mid(strSerials, 2)
        End If
        
        If gblnXWLog = True Then
            Call WriteCommLog("DViewer_DblClick", "调用OpenViewer接口", "序列参数为：" & strSerials)
        End If
    
        Call OpenViewer(1, Nothing, mobjStudyInfo.lngAdviceId, False, Me, strSerials)
        
        Exit Sub
    Else
        '图像在中联FTP管理
        If gblnUseXinWangView = True Then
            '如果是老版本的数据，且使用了新网观片系统，则直接传递远程目录文件名
        
            If lvwSeq.SelectedItem Is Nothing Then Exit Sub '当前检查没有图像，就退出
            
            Set rsTemp = GetStudyImageData(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)

            strImages = ""
            For Each Item In lvwSeq.ListItems
                strSeqUID = Mid(Item.Key, 2)
                If Item.Checked Then
                    '只有当前序列被勾选了，而且选择可部分图象或者全部图象，才打开该序列
                    rsTemp.Filter = "序列UID='" & strSeqUID & "'"
                    While Not rsTemp.EOF
                        If nvl(rsTemp!设备号1) <> "" Then
                            strFtpURL = "\\" & nvl(rsTemp!Host1) & "\" & gstrImageShareDir & nvl(rsTemp!Root1) & nvl(rsTemp!Url)
                        Else
                            strFtpURL = "\\" & nvl(rsTemp!Host2) & "\" & gstrImageShareDir & nvl(rsTemp!Root2) & nvl(rsTemp!Url)
                        End If

                        If strImages <> "" Then strImages = strImages & "[;]"

                        strFtpURL = Replace(strFtpURL, "//", "/")
                        strImages = strImages & Replace(strFtpURL, "/", "\")

                        rsTemp.MoveNext
                    Wend
                End If
            Next

            '打开远程目录文件进行对比观片
            Call OEMViewOpen(0, strImages, 0, mobjStudyInfo.strImgType)
            
            Exit Sub
        End If
    End If
    
    '--------------------上面程序执行完会Exit Sub
    
    '图像在中联数据库，使用观片站打开图像
    '判断是否打开当前图像，如果当前检查没有图像，则打开最近一次检查的图像
    If lvwSeq.SelectedItem Is Nothing Then
        Call OpenLatestImage(Me, mobjPacsCore, mobjStudyInfo, mintViewHistoryImageDays)
    Else
        '规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象
        strImages = ""
        strSerials = ""
        For Each Item In lvwSeq.ListItems
            strSeqUID = Mid(Item.Key, 2)
            If Item.Checked Then
                '只有当前序列被勾选了，而且选择可部分图象或者全部图象，才打开该序列
                If Item.SubItems(1) <> "" Then          '为空表示没有选择任何图象
                    strSerials = strSerials & ",'" & strSeqUID & "'"
                    If strImages = "" Then
                        strImages = strSeqUID & "|" & Item.SubItems(1)
                    Else
                        strImages = strImages & "+" & strSeqUID & "|" & Item.SubItems(1)
                    End If
                End If
            End If
        Next
        If Len(strSerials) = 0 Then         '没有选择任何序列,则默认打开当前序列的图象
            strSerials = ",'" & Mid(lvwSeq.SelectedItem.Key, 2) & "'"
            If lvwSeq.SelectedItem.SubItems(1) <> "" Then
                strImages = Mid(lvwSeq.SelectedItem.Key, 2) & "|" & lvwSeq.SelectedItem.SubItems(1)
            Else
                strImages = Mid(lvwSeq.SelectedItem.Key, 2) & "|全部"
            End If
        End If
        
        strSerials = Mid(strSerials, 2)
        
        intImageInverval = Val(Me.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)
        
        OpenViewer 1, mobjPacsCore, mobjStudyInfo.lngAdviceId, mblnAddImage, Me, strSerials, mobjStudyInfo.blnMoved, intImageInverval, strImages
    End If
    Exit Sub
CallError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    
    With DViewer
        i = .ImageIndex(X, Y)
        If i > 0 And i <= .Images.Count And i <> iCurImageIndex Then
            .Images(iCurImageIndex).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            iCurImageIndex = i
        End If
    End With
End Sub

Private Sub Form_Activate()
On Error GoTo errhandle
    If Me.tag = "Loading" Then Me.tag = ""
        
errhandle:
End Sub

Private Sub Form_Load()
    Dim objFileSystem As New Scripting.FileSystemObject
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim Pane1 As Pane
    Dim strRegPath As String
    
    '读取本地参数
    strRegPath = "公共模块\" & App.ProductName & "\frmPacsImg"
    mintSelectAllSeq = Val(GetSetting("ZLSOFT", strRegPath, "SelectAllSeq", 0))
    mintSelectAllImg = Val(GetSetting("ZLSOFT", strRegPath, "SelectAllImg", 0))
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOfficeXP
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    With Me.cbrMain.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        '.SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "影像显示")
            cbrControl.iconid = 3061: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "显示当前序列影像缩略图"
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "全选序列")
            cbrControl.iconid = 3010: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "选中当前所有序列"
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "全清序列")
            cbrControl.iconid = 3004: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "清除选中当前所有序列"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_SelectAllImages, "全选图像")
            cbrControl.iconid = 227: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "选中当前所有图像"
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_UnSelectAllImages, "全清图像")
        cbrControl.iconid = 229: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "清除选中当前所有图像"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReverseSelectImages, "反选图像")
        cbrControl.iconid = 3012: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "反向选择所有图像"
        Set cbrControl = .Add(xtpControlComboBox, conMenu_Manage_ImageInterval, "图像间隔")
            cbrControl.ToolTipText = "设置打开图像时，图像之间的间隔数量"
            cbrControl.AddItem "0"
            cbrControl.AddItem "2"
            cbrControl.AddItem "3"
            cbrControl.AddItem "4"
            cbrControl.AddItem "5"
            cbrControl.AddItem "7"
            cbrControl.AddItem "10"
            cbrControl.ListIndex = 0
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
            cbrControl.iconid = 791: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "刷新当前病人图像序列": cbrControl.flags = xtpFlagRightAlign
    End With
        
    Call subSetMenuState
    
    '判断当前用户是否具有 观片站的基本权限
    mblnObserve = CheckPopedom(";" & GetPrivFunc(glngSys, 1289) & ";", "基本")

    With dkpMain
        .SetCommandBars Me.cbrMain
        .options.UseSplitterTracker = False '实时拖动
        .options.ThemedFloatingFrames = False
        .options.AlphaDockingContext = True
        .options.HideClient = True
        Set Pane1 = .CreatePane(1, 0, 300, DockTopOf, Nothing)
            Pane1.Handle = lvwSeq.hwnd
            Pane1.options = PaneNoCaption Or PaneNoCloseable
            
        Set Pane1 = .CreatePane(2, 0, 300, DockBottomOf, Pane1)
            Pane1.Handle = lvwImage.hwnd
            Pane1.options = PaneNoCaption Or PaneNoCloseable
            
        Set Pane1 = .CreatePane(3, 0, 400, DockBottomOf, Nothing)
            Pane1.Handle = picView.hwnd
            Pane1.options = PaneNoCaption Or PaneNoCloseable
    End With
    
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    Call RestoreWinState(Me, App.ProductName)
    
'    gblnUseXinWangView = IIf(RegOpenKey(HKEY_CURRENT_USER, "\Software\Silver\Silver Pacs", lngKey) = 0&, True, False) 'IIf(InStr(GetPrivFunc(glngSys, G_LNG_XWPACSVIEW_MODULE), "基本") > 0, True, False)
    
   '如果是RIS工作站，则连接新网数据库，读取参数
    If gblnUseXinWangView Then
        '    挂上截获消息的hook
'        plngXWPreWndProc = XWHook(mobjOwner.hWnd)
        
        Call XWDBServerOpen
        
        mblnAutoOpenViewer = (Val(zlDatabase.GetPara("XW自动打开观片站", glngSys, G_LNG_XWPACSVIEW_MODULE, 1)) = 1)
        If mblnAutoOpenViewer = True Then
            Call XWADViewerStart
        End If
    End If
End Sub

Private Sub ShowSeqList()
'-----------------------------------------------------------------------------------------
'功能：查询检查序列
'参数：无
'返回：无
'-----------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim strCurKey As String
    
    On Error GoTo DBError
    If Not lvwSeq.SelectedItem Is Nothing Then strCurKey = lvwSeq.SelectedItem.Key
    
    With lvwSeq
        
        If .ColumnHeaders.Count <> 7 Then
            With .ColumnHeaders
                .Clear
                .Add , , "影像类别", 2000
                .Add , , "打开图像", 2000
                .Add , , "检查号", 1000
                .Add , , "序列号", 1000
                .Add , , "图像数", 1000
                .Add , , "说明", 2500
                .Add , , "采集时间", 2500
            End With
            .ListItems.Add , , "Temp"
        End If
        
        .ListItems.Clear
    End With
    
    strSQL = "Select A.序列UID,A.序列号,A.序列描述,A.采集时间,B.影像类别,B.检查号," & _
        " B.检查UID,Sum(1) As 图像数 " & _
        "From 影像检查序列 A,影像检查记录 B,影像检查图象 D " & _
        "Where A.检查UID=B.检查UID  And A.序列UID=D.序列UID And B.医嘱ID= [1]  And B.发送号= [2] " & _
        "Group By A.序列UID,A.序列号,A.序列描述,A.采集时间,B.影像类别,B.检查号,B.检查UID " & _
        "Order By B.影像类别,B.检查号,A.序列号"
        
    If mobjStudyInfo.blnMoved Then
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
    End If
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo)
   
    lvwSeq.tag = ""
    If Not rsTmp.EOF Then
        lvwSeq.tag = nvl(rsTmp("检查UID"))
        Do While Not rsTmp.EOF
            
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTmp("序列UID"), rsTmp("影像类别"))
            With tmpItem
                If mintSelectAllImg = 0 Or mintSelectAllImg = 1 Then
                    .SubItems(1) = "全部"
                Else
                    .SubItems(1) = ""
                End If
                
                .SubItems(2) = nvl(rsTmp("检查号"))
                .SubItems(3) = nvl(rsTmp("序列号"))
                .SubItems(4) = nvl(rsTmp("图像数"), 0)
                .SubItems(5) = nvl(rsTmp("序列描述"))
                .SubItems(6) = nvl(rsTmp("采集时间"), date)
                
                If .Key = strCurKey Then .Selected = True
            End With
            rsTmp.MoveNext
        Loop
    End If

    If lvwSeq.Sorted = True Then
        Call lvwSeqSort(lvwSeq.SortKey)
    End If
    
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowImageList(ByVal Item As MSComctlLib.ListItem)
'-----------------------------------------------------------------------------------------
'功能：查询检查序列
'参数：无
'返回：无
'-----------------------------------------------------------------------------------------
    Dim strSeriesUID As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim strCurKey As String
    Dim strOpenImages As String
    Dim ImagesArray() As String
    Dim iSegment As Integer
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iSegCount As Integer
    
    If Not lvwImage.SelectedItem Is Nothing Then strCurKey = lvwImage.SelectedItem.Key
    With lvwImage
        With .ColumnHeaders
            .Clear
            .Add , , "图像号", 2000
            .Add , , "图像描述", 6000
        End With
        .ListItems.Add , , "Temp"
        .ListItems.Clear
    End With
    
    If Item Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo err
    strOpenImages = Item.SubItems(1)
    If strOpenImages <> "全部" And strOpenImages <> "" Then
        ImagesArray = Split(strOpenImages, ";")
        iSegment = 0
        iSegCount = UBound(ImagesArray)
        iStart = Split(ImagesArray(iSegment), "-")(0)
        iEnd = Split(ImagesArray(iSegment), "-")(1)
    End If
    strSeriesUID = Mid(Item.Key, 2)
    strSQL = "Select 图像号,图像描述,图像UID From 影像检查图象 Where 序列UID = [1] Order By 图像号"
    If mobjStudyInfo.blnMoved Then
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取图像信息", strSeriesUID)
    
    lvwImage.tag = ""
    If Not rsTmp.EOF Then
        lvwImage.tag = strSeriesUID
        Do While Not rsTmp.EOF
            Set tmpItem = lvwImage.ListItems.Add(, rsTmp("图像UID"), rsTmp("图像号"))
            With tmpItem
                .SubItems(1) = nvl(rsTmp("图像描述"))
                If strOpenImages = "全部" Then
                    tmpItem.Checked = True
                ElseIf strOpenImages = "" Then
                    tmpItem.Checked = False
                Else
                    If rsTmp("图像号") >= iStart And rsTmp("图像号") <= iEnd Then
                        '满足条件，是需要选中的
                        tmpItem.Checked = True
                    ElseIf rsTmp("图像号") > iEnd Then
                        '大于本段终止号码，则段号加1 ，重新调整起始号码和终止号码
                        iSegment = iSegment + 1
                        If iSegment > iSegCount Then
                            tmpItem.Checked = False
                        Else
                            iStart = Split(ImagesArray(iSegment), "-")(0)
                            iEnd = Split(ImagesArray(iSegment), "-")(1)
                            If rsTmp("图像号") >= iStart And rsTmp("图像号") <= iEnd Then
                                tmpItem.Checked = True
                            Else
                                tmpItem.Checked = False
                            End If
                        End If
                    Else
                        '小于本段起始号码，则不选中
                        tmpItem.Checked = False
                    End If
                End If
                If .Key = strCurKey Then .Selected = True
            End With
            rsTmp.MoveNext
        Loop
    End If
    
    DViewer.Images.Clear: iCurImageIndex = 0
    
    If lvwImage.ListItems.Count >= 1 Then
        Call ShowLvwImage(lvwImage.ListItems(1))
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    Set mObjOwner = Nothing
    Set mObjActiveMenuBar = Nothing
    Set mobjPacsCore = Nothing
    
    strRegPath = "公共模块\" & App.ProductName & "\frmPacsImg"
    SaveSetting "ZLSOFT", strRegPath, "SelectAllSeq", mintSelectAllSeq
    SaveSetting "ZLSOFT", strRegPath, "SelectAllImg", mintSelectAllImg
    
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveWinState(Me, App.ProductName)
    
    '如果是RIS工作站，则断开跟新网数据库的连接
    If gblnUseXinWangView Then
        '    卸载hook
'        XWUnhook mobjOwner.hWnd, plngXWPreWndProc
        
        Call XWDBServerClose
        Call XWADViewerExit
    End If
End Sub

Private Sub lvwImage_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call WriteSelectdImages(lvwImage.tag)
End Sub

Private Sub lvwImage_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If Item.Checked <> Item.Selected Then
        Item.Checked = Item.Selected
        Call WriteSelectdImages(lvwImage.tag)
    End If
    Call ShowLvwImage(Item)
End Sub

Private Sub ShowLvwImage(ByVal Item As MSComctlLib.ListItem)
    Dim strImageUID As String
    
    If mblnShowPic = False Then
        DViewer.Images.Clear
        Exit Sub
    End If
    
    On Error GoTo DBError
    strImageUID = Item.Key
    '读取图像到DViewer中
    GetAllImages Me, DViewer, mobjStudyInfo.blnMoved, 3, 0, lvwImage.tag, 1, 1, False, "", strImageUID

    If DViewer.Images.Count > 0 Then
        iCurImageIndex = 1
    Else
        iCurImageIndex = 0
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lvwImage.ListItems.Count >= 1 And Button = 2 Then Call ShowPopupImage(False)
End Sub

Private Sub lvwSeq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     Call lvwSeqSort(ColumnHeader.Index - 1)
End Sub

Private Sub lvwSeqSort(intSortKey As Integer)
    Dim i As Integer
    
    lvwSeq.Sorted = False
    lvwSeq.SortKey = intSortKey
    lvwSeq.SortOrder = IIf(lvwSeq.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    
    '对数值型的数据排序
    If intSortKey = 3 Or intSortKey = 4 Then
        For i = 1 To lvwSeq.ListItems.Count
            lvwSeq.ListItems(i).SubItems(intSortKey) = Format(lvwSeq.ListItems(i).SubItems(intSortKey), "0000000000")
        Next i
        lvwSeq.Sorted = True
        For i = 1 To lvwSeq.ListItems.Count
            lvwSeq.ListItems(i).SubItems(intSortKey) = Val(lvwSeq.ListItems(i).SubItems(intSortKey))
        Next i
    Else
        lvwSeq.Sorted = True
    End If
End Sub

Private Sub lvwSeq_DblClick()
    If Not mblnObserve Then Exit Sub
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
    DViewer_DblClick
End Sub

Private Sub lvwSeq_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '图像在中联PACS，才支持对序列的选择
    If mobjStudyInfo.intImageLocation = 0 Then
        lvwSeq.SelectedItem = Item
        Call ShowImageList(Item)
    End If
End Sub

Private Sub lvwSeq_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '图像在中联PACS，才支持对序列的选择
On Error GoTo errhandle
    If mobjStudyInfo.intImageLocation = 0 Then
        If Item.Checked <> Item.Selected Then
            Item.Checked = Item.Selected
        End If
        Call ShowImageList(Item)
    Else
        mlngCurImageCount = Item.SubItems(3)
    End If
        
    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub lvwSeq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ''图像在中联PACS，才支持删除序列的弹出菜单
    If mobjStudyInfo.intImageLocation = 0 And lvwSeq.ListItems.Count >= 1 And Button = 2 Then
        Call ShowPopupImage(True)
    End If
End Sub

Private Sub picView_Resize()
On Error GoTo errhandle
    Dim iCols As Integer, iRows As Integer
    
    With DViewer
        .Left = 0: .Top = 0
        .Width = picView.ScaleWidth: .Height = picView.ScaleHeight
        
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
errhandle:
End Sub

Public Function ZLfun3DImgProcess(blnAutoDecompress As Boolean) As String
'------------------------------------------------
'功能：三维重建预处理，移动当前被选中序列的图像
'参数： blnAutoDecompress -- True,下载后解压缩，False ，直接下载不作处理
'返回：图像被移动的目的目录，如果移动失败则返回空
'------------------------------------------------

    Dim strSeriesUID As String
    Dim Item As MSComctlLib.ListItem
    Dim iSeriesCount As Integer
    
    On Error GoTo CallError
    If lvwSeq.SelectedItem Is Nothing Then
        MsgBoxD Me, "请选择一个序列进行三维重建。", vbInformation, Me.Caption
        ZLfun3DImgProcess = ""
        Exit Function
    End If
    
    iSeriesCount = 0
    For Each Item In lvwSeq.ListItems
        If Item.Checked Then
            iSeriesCount = iSeriesCount + 1
            strSeriesUID = Mid(Item.Key, 2)
        End If
    Next
    
    '判断是否只有多个序列被选择，三维重建一次只能处理一个序列
    If iSeriesCount <> 1 Then
        MsgBoxD Me, "请选择一个序列进行三维重建，每次重建只能选择一个系列。", vbInformation, Me.Caption
        ZLfun3DImgProcess = ""
        Exit Function
    End If
    
    '移动指定序列UID的图像
    ZLfun3DImgProcess = funMove3DImage(strSeriesUID, mobjStudyInfo.blnMoved, blnAutoDecompress)
    Exit Function
CallError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ZLfun3DImgProcess = ""
End Function

Private Function funMove3DImage(strSeriesUID As String, blnMoved As Boolean, blnDecompress As Boolean) As String
'------------------------------------------------
'功能：将一个序列的图像移动到3D临时目录中，等待三维重建软件的调用
'参数：
'       strSeriesUID -- 图像的序列UID
'       blnMoved -- 图像是否被转储
'       blnDecompress -- 下载图像后是否解压缩，True，解压缩，False，下载后不作处理
'返回：图像被移动的目的目录，如果移动失败则返回空
'------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim struFtpTag As TFtpConTag
    Dim lngResult As Long
    
    Dim str3DCachePath As String
    Dim strTmpFile As String
    Dim strImageFullPath As String
    Dim dcmImages As New DicomImages
    
    strSQL = "Select A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1," & _
        "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As 图像目录,A.图像UID,d.设备号 as 设备号1, " & _
        "E.FTP用户名 As User2,E.FTP密码 As Pwd2," & _
        "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2," & _
        "e.设备号 as 设备号2,C.检查UID,B.序列UID " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) "
    If blnMoved Then
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If

    On Error GoTo DBError
    strSQL = strSQL & "And A.序列UID= [1] Order By A.图像号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取图像", strSeriesUID)
    
    If rsTmp.RecordCount > 0 Then
        
        '创建本地目录,3D图像目录由前缀"App.Path & "\TmpImage\3D"+接收日期+检查UID+序列UID
        str3DCachePath = FormatFilePath(GetAppRootPath() & "\Apply\TmpImage\3D\" & Replace(nvl(rsTmp("图像目录")), "/", "\") & "\" & strSeriesUID & "\")
        strImageFullPath = FormatFilePath(GetAppRootPath() & "\Apply\TmpImage\" & Replace(nvl(rsTmp("图像目录")), "/", "\") & "\")
        MkLocalDir str3DCachePath

        On Error GoTo DBError
        
        Do While Not rsTmp.EOF
            '如果3D目录下没有图像，再检查本地缓存目录，最后再从FTP下载图像
            If blnDecompress Then
                '如果自动解压缩，则本地图像目录文件名需要修改
                strTmpFile = str3DCachePath & "3DTemp"
            Else
                strTmpFile = str3DCachePath & nvl(rsTmp("图像UID"))
            End If
            
            If Dir(strTmpFile) = vbNullString Then  '有图像则不需要做任何操作
                If Dir(strImageFullPath & nvl(rsTmp("图像UID"))) = vbNullString Then
                    '本地缓存图像不存在，则读取FTP图像
                    '建立FTP连接
                    struFtpTag = FtpTagInstance(nvl(rsTmp("Host1")), _
                                                nvl(rsTmp("User1")), _
                                                nvl(rsTmp("Pwd1")), _
                                                nvl(rsTmp("Root1")) & nvl(rsTmp("图像目录")))
                    
                    If Trim(struFtpTag.Ip) = "" Then
                        struFtpTag = FtpTagInstance(nvl(rsTmp("Host2")), _
                                                    nvl(rsTmp("User2")), _
                                                    nvl(rsTmp("Pwd2")), _
                                                    nvl(rsTmp("Root2")) & nvl(rsTmp("图像目录")))
                    End If
                    
                    lngResult = FtpDownload(struFtpTag, nvl(rsTmp!图像UID), strTmpFile, False)
                    If lngResult = frAbort Then Exit Function

                Else
                '本地观片缓存中图像存在，直接复制到3D目录
                    FileCopy strImageFullPath & nvl(rsTmp("图像UID")), strTmpFile
                End If
                
                '如果自动解压缩，则打开已经下载好的临时文件，解压缩后再保存
                If blnDecompress Then
                    dcmImages.ReadFile strTmpFile
                    dcmImages(1).WriteFile str3DCachePath & nvl(rsTmp("图像UID")), True
                    dcmImages.Clear
                    Kill strTmpFile
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    funMove3DImage = str3DCachePath
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
    funMove3DImage = ""
End Function

Private Sub ShowSeqImg()
On Error GoTo err
    '根据图像所在的PACS位置，调用不同的过程来显示序列列表
    If mobjStudyInfo.intImageLocation = 1 Or mobjStudyInfo.intImageLocation = 2 Then  '图像在新网PACS或者云平台
    
        If dkpMain.Panes(3).Closed = False Then dkpMain.Panes(3).Closed = True
        If dkpMain.Panes(2).Closed = False Then dkpMain.Panes(2).Closed = True
        
'        lvwImage.Visible = False
'        lvwSeq.Visible = False
        
        Call showXWSeq

        lvwImage.ListItems.Clear
        lvwImage.ColumnHeaders.Clear
        
        DViewer.Images.Clear
    Else
    
        If dkpMain.Panes(3).Closed Then dkpMain.Panes(3).Closed = False
        If dkpMain.Panes(2).Closed Then dkpMain.Panes(2).Closed = False
        
        Call ShowSeqList     '显示序列
        
        If lvwSeq.SelectedItem Is Nothing Then
            DViewer.Images.Clear
            Call ShowImageList(Nothing)
        ElseIf mintSelectAllSeq = 0 Then
            lvwSeq_ItemClick lvwSeq.SelectedItem
        ElseIf mintSelectAllSeq = 1 Then
            SelectAllSeq True
        ElseIf mintSelectAllSeq = 2 Then
            SelectAllSeq False
        End If
        
        If lvwImage.SelectedItem Is Nothing Then
            DViewer.Images.Clear
        Else
            ShowLvwImage lvwImage.SelectedItem
        End If
    End If
    
    lvwImage.Enabled = IIf(mobjStudyInfo.intImageLocation = 0, True, False)
    lvwImage.HideColumnHeaders = IIf(mobjStudyInfo.intImageLocation = 0, False, True)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub WriteSelectdImages(strSeriesUID As String)
    Dim i As Integer
    Dim j As Integer
    Dim strOpenImages As String
    Dim blnSelectAll As Boolean
    Dim blnSelectNone As Boolean
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iSegment As Integer
    
    blnSelectNone = True
    blnSelectAll = True
    For j = 1 To lvwImage.ListItems.Count
        If lvwImage.ListItems(j).Checked = True Then
            blnSelectNone = False
            '开始记录本段
            If iStart <> 0 Then
                iEnd = lvwImage.ListItems(j).Text
            Else
                iStart = lvwImage.ListItems(j).Text
                iEnd = lvwImage.ListItems(j).Text
            End If
        Else
            blnSelectAll = False
            '结束记录本段
            If iStart <> 0 Then
                iSegment = iSegment + 1
                If strOpenImages = "" Then
                    strOpenImages = iStart & "-" & iEnd
                Else
                    strOpenImages = strOpenImages & ";" & iStart & "-" & iEnd
                End If
                iStart = 0
                iEnd = 0
            End If
        End If
    Next j
    If iStart <> 0 Then
        iSegment = iSegment + 1
        If strOpenImages = "" Then
            strOpenImages = iStart & "-" & iEnd
        Else
            strOpenImages = strOpenImages & ";" & iStart & "-" & iEnd
        End If
    End If
    If blnSelectAll = True Then
        strOpenImages = "全部"
    End If
    If blnSelectNone = True Then
        strOpenImages = ""
    End If
    
    For i = 1 To lvwSeq.ListItems.Count
        If lvwSeq.ListItems(i).Key = "_" & strSeriesUID Then
            lvwSeq.ListItems(i).ListSubItems(1) = strOpenImages
        End If
    Next i
End Sub

Private Sub ShowPopupImage(blnIsSeries As Boolean)
'------------------------------------------------
'功能：创建鼠标右键弹出菜单
'参数： blnIsSeries -- True 序列菜单；False 图像菜单
'------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    If Not CheckPopedom(mstrPrivs, "清除图像") Then Exit Sub
    If mobjStudyInfo.blnMoved Then Exit Sub
    If mblnIsHistoryMode Then Exit Sub
    
    '鼠标右键弹出菜单
    Set cbrToolBar = cbrMain.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        If blnIsSeries Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_DeleteSeries, "删除序列")
         Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_DeleteImage, "删除图像")
         End If
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub zlMenuDeleteImageClick(lngControlID As Long)
'------------------------------------------------
'功能：删除当前选中的图像
'参数： lngControlID -- 按钮ID
'功能：删除图像
'------------------------------------------------
    Dim i As Integer
    Dim blImgDeleted As Boolean '是否有图像被删除--true 是
    Dim blnIsCancel As Boolean
    
    On Error GoTo err
    blImgDeleted = False
    
    If MsgBoxD(Me, "确定要删除所有勾选中的图像吗？", vbOKCancel, "删除图像") = vbCancel Then Exit Sub
    
    If Not mObjNotify Is Nothing Then
        Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 0, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, blnIsCancel)
        If blnIsCancel Then Exit Sub
    End If
    
    If lngControlID = conMenu_Manage_DeleteImage Then
        '删除当前勾选的图像
        For i = 1 To lvwImage.ListItems.Count
            If lvwImage.ListItems(i).Checked = True Then
                If DeleteImages(Me, 1, lvwImage.ListItems(i).Key, "") = False Then
                    If i = 1 Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
                
                blImgDeleted = True
            End If
        Next
    ElseIf lngControlID = conMenu_Manage_DeleteSeries Then
        '删除当前勾选的序列
        For i = 1 To lvwSeq.ListItems.Count
            If lvwSeq.ListItems(i).Checked = True Then
                If DeleteImages(Me, 2, "", Mid(lvwSeq.ListItems(i).Key, 2)) = False Then
                    If i = 1 Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
                
                blImgDeleted = True
            End If
        Next
    End If
        
    '刷新列表显示
    Call ShowSeqImg
    
    '如果最后一个序列也被删除了,应该刷新列表
    If lvwSeq.ListItems.Count = 0 Then
        If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, -1, hwnd)
    Else
        If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, , hwnd)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub showXWSeq()
'------------------------------------------------
'功能：显示新网PACS中图像序列
'参数：无
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim lngImgCount As Long
    
    On Error GoTo err
    
    If gcnXWDBServer.State <> adStateOpen Then Exit Sub
    
    With lvwSeq
        If .ColumnHeaders.Count <> 6 Then
            With .ColumnHeaders
                .Clear
                .Add , , "影像类别", 2000
                .Add , , "检查号", 1000
                .Add , , "序列号", 1000
                .Add , , "图像数", 900
                .Add , , "说明", 4000
                .Add , , "采集时间", 2500
            End With
            .ListItems.Add , , "Temp"
        End If
        .ListItems.Clear
    End With
    
    strSQL = "select F_SER_ID as SERIES主键,F_STU_ID as Study主键,F_SER_NO as 序列号,F_COUNT_IMG as 图像数, F_SER_DATE as 序列日期,F_SER_TIME as 序列时间, " _
                & " F_SER_CONTEXT as 序列描述,F_MODALITY as 影像类型,F_STU_NO as 医嘱ID from V_OEM_SERIES where F_STU_NO ='" & mobjStudyInfo.lngAdviceId & "' order by F_SER_NO"
    Set rsTemp = gcnXWDBServer.Execute(strSQL)
    
    lngImgCount = 0
    lvwSeq.tag = ""
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            lngImgCount = lngImgCount + nvl(rsTemp!图像数, 0)
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTemp!SERIES主键, rsTemp!影像类型)
            With tmpItem
                .SubItems(1) = nvl(rsTemp!Study主键)
                .SubItems(2) = nvl(rsTemp!序列号)
                .SubItems(3) = nvl(rsTemp!图像数)
                .SubItems(4) = nvl(rsTemp!序列描述)
                .SubItems(5) = Replace(nvl(rsTemp!序列日期, date), ".", "-") + " " + nvl(rsTemp!序列时间, time)
                .Checked = True
            End With
            rsTemp.MoveNext
        Loop
    End If
        
    Exit Sub
    
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

