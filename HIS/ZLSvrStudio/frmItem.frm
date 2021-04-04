VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmItem 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TaskPanel tplFunc 
      Height          =   4770
      Left            =   30
      TabIndex        =   0
      Top             =   315
      Width           =   3210
      _Version        =   589884
      _ExtentX        =   5662
      _ExtentY        =   8414
      _StockProps     =   64
      Behaviour       =   1
      ItemLayout      =   2
      HotTrackStyle   =   3
   End
   Begin XtremeSuiteControls.ShortcutCaption stcItem 
      Height          =   300
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   3210
      _Version        =   589884
      _ExtentX        =   5662
      _ExtentY        =   529
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==变量定义
'==============================================================
Public gstrParentNo     As String
Public gstrParentCap    As String
Private mstrDefNo       As String '缺省功能

'==============================================================
'==公共接口
'==============================================================
Public Function RunByModule(Optional ByVal strModule As String) As Boolean
'功能：执行某个模块功能，当不传时，执行默认模块功能
    Dim frmChild As Form
    Dim GroupTmp As TaskPanelGroup
    Dim itemTmp As TaskPanelGroupItem
    Dim strTmp As String
    
    If strModule = "" Then
        strModule = mstrDefNo
    ElseIf strModule = frmMDIMain.gstrLastModule Then
        Exit Function
    End If
    
    For Each frmChild In Forms
        If Not gfrmActive Is Nothing Then
            If gfrmActive.name <> frmChild.name And InStr(",frmMDIMain,frmPubIcons,frmItem,", "," & frmChild.name & ",") <= 0 Then
                Unload frmChild
            End If
        End If
    Next
    tplFunc.Tag = "直接设置"
    For Each GroupTmp In tplFunc.Groups
        For Each itemTmp In GroupTmp.Items
            If Val(strModule) = itemTmp.Id Then
                itemTmp.Selected = True
            Else
                itemTmp.Selected = False
            End If
        Next
    Next
    tplFunc.Tag = ""
    frmMDIMain.gstrLastModule = strModule
    If Not gfrmActive Is Nothing Then
        Unload gfrmActive
        Set gfrmActive = Nothing
    End If
    
    'DBA工具只有DBA可以使用,在这里判断
    If strModule Like "06*" Then
            If Not gblnDBA Then
                frmMDIMain.stbThis.Panels(2).Text = "当前用户不是DBA用户，权限不足，无法使用该功能。"
                Exit Function
            End If
    End If
    
    Select Case strModule
        Case "01", "02", "03", "04", "05" '装卸管理
            frmDescribe.mstr编号 = strModule
            Set gfrmActive = frmDescribe
        Case "0101" '系统装卸管理
            Set gfrmActive = frmAppStart
        Case "0102" '系统升迁
            If Not CheckAndAdjustMustTable("ZLRegInfo", , True) Then
                Exit Function
            End If
            If Not CheckAndAdjustMustTable("zlUpgradeConfig", , True) Then
                Exit Function
            End If
            Set gfrmActive = frmAppUpgrade
        Case "0103" '对象检查修复
            Set gfrmActive = frmAppCheck
        Case "0105" '编译无效对象
            Set gfrmActive = frmCompileInvalid
        Case "0104" '置换安装脚本
            Set gfrmActive = frmAppScript
        Case "0201" '刘兴宏:历史数据空间管理
            Set gfrmActive = frmHistoryDataMgr 'frmDataMove
        Case "0202" '数据导出
            Set gfrmActive = frmExp
        Case "0203" '数据导入
            Set gfrmActive = frmImp
        Case "0204" '数据调出
            Set gfrmActive = frmLoadOut
        Case "0205" '数据调入
            Set gfrmActive = frmLoadIn
        Case "0206" '数据清除
            If MsgBox("该程序运行需要等待一段时间，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
            Set gfrmActive = frmClearData
        Case "0207" '数据连接
            Set gfrmActive = frmConnManagerParent
        Case "0208" '检验图片数据转移
            If Not gblnSystemUser Then
                frmMDIMain.stbThis.Panels(2).Text = "当前用户不是标准版系统所有者，无法使用该功能"
                Exit Function
            End If
            Set gfrmActive = frmLisPic2Ftp
        Case "0301" '用户注册管理
            Set gfrmActive = frmRegist
        Case "0302" '运行状态监控
            Set gfrmActive = frmStatus
        Case "0303" '后台作业管理
            Set gfrmActive = frmAutoJobs
        Case "0304" '运行日志管理
            Set gfrmActive = FrmRunLog
        Case "0305" '错误日志管理
            Set gfrmActive = FrmErrLog
        Case "0306" '系统运行选项
            Set gfrmActive = FrmRunOption
        Case "0307" '客户端升级管理
            Set gfrmActive = frmClientUpgradeManage
        Case "0308" '站点运行控制
            Set gfrmActive = frmClientsParas
        Case "0310" '系统参数管理
            Set gfrmActive = frmParameters
        Case "0312" '医院信息维护
            Set gfrmActive = frmUnitInfoEdit
        Case "0314" '重要操作变动日志管理
            Set gfrmActive = frmAuditLogManage
        Case "0315" '功能限时管理
            Set gfrmActive = frmRunLimitManage
        Case "0401" '角色授权管理
            Set gfrmActive = frmRole
        Case "0402" '用户授权管理
            Set gfrmActive = frmUser
        Case "0403" '菜单重组规划
            Set gfrmActive = frmMenu
        Case "0404" '管理工具授权
            Set gfrmActive = frmMgrGrant
        Case "0501", "0502", "0505" '报表管理
            frmRptMan.mstr编号 = strModule
            Set gfrmActive = frmRptMan
        Case "0503"
            Set gfrmActive = frmInputTools
        Case "0504"
            Set gfrmActive = frmNoticeTools
        Case "0601", "0602", "0606", "0604", "0605" 'DBA管理工具
            Set gfrmActive = frmDbatoolsParent
            strTmp = strModule
        Case "0603"  'SQL跟踪工具
            Set gfrmActive = frmSQLTrace
        Case "0607"     '用户与IP限制
            Set gfrmActive = frmUserLimit
        Case "0608"     '应用程序授权
            Set gfrmActive = frmAppLimit
        Case "0609"     '用户登录日志
            Set gfrmActive = frmLoginLog
        Case "0610"     '对象审计管理
            Set gfrmActive = frmFga
        Case ""
    End Select
    If Not gfrmActive Is Nothing Then
        frmMDIMain.stbThis.Panels(2).Text = ""
        Call FindWindowAndSetActive(gfrmActive)
        
        If gfrmActive.name = "frmDbatoolsParent" Then
            gfrmActive.ShowToolsForm strTmp
        Else
            gfrmActive.Show
        End If
        gfrmActive.ZOrder 0
    End If
    RunByModule = True
End Function

'==============================================================
'=控件事件
'==============================================================
Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    On Error GoTo errH
    
    With frmMDIMain.grsToolsMenu
        If Not frmMDIMain.grsToolsMenu Is Nothing Then
            .Filter = "上级=" & IIf(gstrParentNo = "", "NULL", "'" & gstrParentNo & "'")
            If .Sort = "" Then
                If CheckAndAdjustMustTable("Zlsvrtools", "次序", False) Then
                    .Sort = "次序,编号"
                Else
                    .Sort = "编号"
                End If
            End If
            Set tpGroup = tplFunc.Groups.Add(Val(gstrParentNo), gstrParentCap)
            Do While Not .EOF
                If mstrDefNo = "" Then mstrDefNo = !编号 & ""
                If !编号 & "" <> "0404" Or gblnSystemUser Then
                    tpGroup.Items.Add(Val(!编号), !标题, xtpTaskItemTypeLink, Val(!编号) + 1).Selected = False
                End If
                .MoveNext
            Loop
            
            tplFunc.SetMargins 1, 2, 0, 2, 2
            tplFunc.SelectItemOnFocus = True
            Call tplFunc.Icons.AddIcons(frmMDIMain.GetIcons.Icons)
            tplFunc.SetIconSize 24, 24
            tpGroup.CaptionVisible = False
            tpGroup.Expanded = True
            stcItem.Caption = gstrParentCap
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    stcItem.Left = Me.Left
    stcItem.Width = Me.Width
    
    tplFunc.Height = Me.Height - Me.stcItem.Height
    tplFunc.Width = Me.Width
    tplFunc.Left = Me.Left
    tplFunc.Top = Me.stcItem.Top + Me.stcItem.Height
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    If tplFunc.Tag = "" Then
        If frmMDIMain.gstrLastModule <> "0" & Item.Id Then
            Call RunByModule("0" & Item.Id)
        End If
    End If
End Sub

'==============================================================
'=私有方法
'==============================================================
Private Sub FindWindowAndSetActive(ByVal FrmObj As Form)
    Dim LngTargetHdl As Long
    '--如果该窗体已经打开,则激活它(这样,窗体的大小不会发生变化)--zyb
    LngTargetHdl = FindWindow(vbNullString, FrmObj.Caption)
    If LngTargetHdl <> 0 Then
        If IsIconic(LngTargetHdl) Then
            Call ShowWindow(LngTargetHdl, 9)            '还原指定窗体为原大小
        End If
        Call SetActiveWindow(LngTargetHdl)
    End If
End Sub



