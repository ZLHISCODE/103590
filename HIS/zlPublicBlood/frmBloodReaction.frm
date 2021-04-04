VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmBloodReaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "输血反应"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11400
   Icon            =   "frmBloodReaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin zlPublicBlood.usrCardEdit UCE 
      Height          =   10725
      Left            =   -30
      TabIndex        =   0
      Top             =   345
      Width           =   10980
      _extentx        =   19368
      _extenty        =   18918
      tabsposition    =   1
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBloodReaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng阶段 As Long      '1：医生处理阶段  2：输血科处理阶段
Private mlng病人ID As Long
Private mlng主页id As Long
Private mlng病人来源 As Long  '1-门诊  2-住院
Private mstrPrivs As String   '权限串
Private mfrmMain As Object    '父窗体
Private mlngSys As Long
Private mlng模块号 As Long
Public mblnBloodReactionIsOpen As Boolean '非模态状态下，判断窗体是否开启

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始化处理
    
    Call CommandBarInit(cbsMain)
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加"): objControl.BeginGroup = True 'objControl.BeginGroup = True就是划竖线
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "提交"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退")
        Set objControl = .Add(xtpControlButton, conMenu_View_Detail, "输血执行"): objControl.ToolTipText = "输血执行情况查阅": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each objControl In objBar.Controls
        If objControl.Type = xtpControlButton Then objControl.Style = xtpButtonIconAndCaption
    Next
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理
    With cbsMain.KeyBindings '
        .Add 0, vbKeyDelete, conMenu_Edit_Delete            '删除
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '新增
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify          '修改
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add FCONTROL, vbKeyS, conMenu_Edit_Save            '保存
        .Add FCONTROL, vbKeyC, conMenu_Edit_Transf_Cancle   '取消
        .Add FCONTROL, vbKeyX, conMenu_File_Exit            '退出
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
    End With
    
    InitCommandBar = True
    Exit Function
ErrHand:
    
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case conMenu_File_Preview '、预览
            Call UCE.ShowPrint(2)
        Case conMenu_File_Print '打印
            Call UCE.ShowPrint(1)
        Case conMenu_Edit_NewItem: '新增
            UCE.AddPage
        Case conMenu_Edit_Modify: '修改
            UCE.ShowModify
        Case conMenu_Edit_Delete: '删除
            If IsPrivs(mstrPrivs, "删除他人") = False Then
                If UCE.Doctor <> "" And UCE.Doctor <> UserInfo.姓名 Then
                    MsgBox "您没有权限删除他人记录的数据！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            UCE.ShowDelete
        Case conMenu_Edit_Save: '保存
            UCE.ShowSave
        Case conMenu_Edit_Transf_Cancle: '取消
            UCE.ShowCancel
        Case conMenu_Edit_Audit: '提交
            UCE.SubmitData
        Case conMenu_Edit_Untread: '回退
            UCE.ShowUntread
        Case conMenu_View_Detail '执行情况查看
            Call frmBloodExecEdit.ViewExecution(Me, UCE.BloodID)
        Case conMenu_Help_Help
            Call gobjComlib.ShowHelp(App.ProductName, Me.hWnd, Me.name, Int((2200) / 100))
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long
    On Error GoTo Errorhand
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    '窗体其它控件Resize处理
    UCE.Move lngLeft, lngTop + 50, lngRight - lngLeft, lngBottom - lngTop
Errorhand:
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case conMenu_File_Preview, conMenu_File_Print
            Control.Visible = IsPrivs(mstrPrivs, "单据打印")
        Case conMenu_Edit_Modify: '修改
            '无记录反应的权限则修改按钮不可见。
            Control.Visible = IsPrivs(mstrPrivs, "记录反应")
            '医生阶段已提交状态 或者 输血科已提交状态 或者 医生阶段输血科新增的数据 或者 新增状态 或者 修改状态 下修改按钮不使能，其他情况使能。
            Control.Enabled = Not ((mlng阶段 <> 2 And UCE.lng状态 <> 0) Or (mlng阶段 = 2 And UCE.lng状态 = 2) Or (mlng阶段 <> 2 And UCE.输血科新增 = True) Or UCE.strST = 新增 Or UCE.strST = 修改)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Delete: '删除
            '无删除记录的权限或者是输血科阶段，则删除按钮不可见。
            Control.Visible = IsPrivs(mstrPrivs, "删除记录")   'and not (mlng阶段=2 and not IsPrivs(mstrPrivs, "输血科新增"))，由于需求，输血科在有相关权限的情况下允许删除
            '医生阶段已提交状态 或者 输血科阶段已提交状态 或者 医生阶段输血科新增的数据 或者 新增状态 或者 修改状态 情况下删除按钮不使能，其他情况使能。
            Control.Enabled = Not ((mlng阶段 <> 2 And UCE.lng状态 <> 0) Or (mlng阶段 = 2 And UCE.lng状态 <> 0) Or (mlng阶段 <> 2 And UCE.输血科新增 = True) Or UCE.strST = 新增 Or UCE.strST = 修改)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_NewItem: '新增
            '输血科没有新增权限时，新增按钮不可见，其他情况新增按钮可见。
            Control.Visible = IsPrivs(mstrPrivs, "记录反应")
            
            '余浪 2017年6月22日  blnAddPage=true也就是当前数据可以修改的情况下，新增按钮不使能，其他情况使能
            Control.Enabled = Not (UCE.blnAddPage = True)
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Save: '保存
            Control.Visible = IsPrivs(mstrPrivs, "记录反应")
            '未提交且数据变化时，保存使能
            Control.Enabled = UCE.DataChanged And UCE.lng状态 = 0
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Transf_Cancle: '取消
            Control.Visible = IsPrivs(mstrPrivs, "记录反应")
            '未提交且数据变化时，取消使能
            Control.Enabled = UCE.DataChanged And UCE.lng状态 = 0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Audit: '提交
            Control.Visible = IsPrivs(mstrPrivs, "提交回退")
            '医生阶段已提交数据 或者 输血科阶段已提交数据 或者 医生阶段输血科新增状态 或者 在新增或修改状态 或者 无病人或者未选中病人时提交不使能，其他状态提交使能。
            Control.Enabled = Not ((mlng阶段 <> 2 And UCE.lng状态 <> 0) Or (mlng阶段 = 2 And UCE.lng状态 = 2) Or (mlng阶段 <> 2 And UCE.输血科新增 = True) Or UCE.strST = 新增 Or UCE.strST = 修改)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Untread: '回退
            Control.Visible = IsPrivs(mstrPrivs, "提交回退")
            '医生阶段非医生已提交状态 或者 输血科阶段未提交状态 或者 医生阶段输血科新增数据 或者 输血科阶段输血科新增数据未提交状态 或者 在新增或修改状态 或者 无病人或者未选中病人 时回退不使能，其他状态回退使能。
            Control.Enabled = Not ((mlng阶段 <> 2 And UCE.lng状态 <> 1) Or (mlng阶段 = 2 And UCE.lng状态 <> 2) Or (mlng阶段 <> 2 And UCE.输血科新增 = True) Or (mlng阶段 = 2 And UCE.lng状态 = 0 And UCE.输血科新增) Or UCE.strST = 新增 Or UCE.strST = 修改)
            
    Case conMenu_View_Detail
        Control.Enabled = UCE.BloodID > 0
    End Select
End Sub

Public Sub BloodReaction(frmMain As Object, lng阶段 As Long, lng病人ID As Long, lng主页id As Long, lng病人来源 As Long, ByVal lngSys As Long, _
                    lng模块号 As Long, Optional strPrivs As String, Optional lngisModul As Long = 0, Optional ByVal lng收发ID As Long = 0)
    '功能：输血反应主要的处理程序
    '参数：lng阶段-医生处理阶段还是输血科处理阶段，lng病人id-病人的id，lng主页id-病人的主页id，lng病人来源-1：门诊、2：住院，strPrivs-权限串；lng收发id-存在则按照收发id产生新增页面
    Set mfrmMain = frmMain
    If mblnBloodReactionIsOpen = True Then GoTo TOSHOW
    If zlGetComLib = False Then MsgBox "获取对象失败！", vbInformation, gstrSysName: Exit Sub
    InitCommandBar
    mlng阶段 = lng阶段 '1：医生处理阶段  2：输血科处理阶段
    mlng病人ID = lng病人ID
    mlng主页id = lng主页id
    mlng病人来源 = lng病人来源 '1：门诊  2：住院
    mstrPrivs = strPrivs
    mlngSys = lngSys
    mlng模块号 = lng模块号
    If zlGetComLib = False Then MsgBox "获取对象失败！", vbInformation, gstrSysName: Exit Sub
    UCE.InitEdit
    UCE.showInfor mlng病人ID, mlng病人来源, mlng主页id, mlng阶段, gcnOracle, Me, mlng模块号, , , lng收发ID
TOSHOW:
    mblnBloodReactionIsOpen = True
    Me.Show lngisModul, mfrmMain '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (UCE.strST = 新增 And UCE.DataChanged = True) Or UCE.strST = 修改 Then
        Cancel = (MsgBox("数据未保存，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    mblnBloodReactionIsOpen = False
End Sub
