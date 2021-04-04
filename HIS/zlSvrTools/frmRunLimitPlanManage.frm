VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmRunLimitPlanManage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "方案设置"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12675
   Icon            =   "frmRunLimitPlanManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.TreeView tvwPlanTree 
      Height          =   5235
      Left            =   45
      TabIndex        =   0
      Top             =   975
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   9234
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img24"
      Appearance      =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPlanDetail 
      Height          =   5115
      Left            =   3825
      TabIndex        =   1
      Top             =   975
      Width           =   8235
      _cx             =   14526
      _cy             =   9022
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16774866
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16774866
      BackColorAlternate=   16774866
      GridColor       =   -2147483633
      GridColorFixed  =   15984570
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRunLimitPlanManage.frx":6852
      ScrollTrack     =   0   'False
      ScrollBars      =   1
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList img24 
      Left            =   2970
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitPlanManage.frx":68E6
            Key             =   "enabled"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitPlanManage.frx":79F8
            Key             =   "enabledLock"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitPlanManage.frx":947A
            Key             =   "disabled"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitPlanManage.frx":AEFC
            Key             =   "disabledLock"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgPlanDetail 
      Left            =   3075
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   97
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitPlanManage.frx":C97E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImageMan 
      Left            =   3765
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmRunLimitPlanManage.frx":10C54
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   3165
      Top             =   135
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Visible         =   0   'False
      Begin VB.Menu mnuViewShow 
         Caption         =   "显示停用方案(&S)"
      End
   End
   Begin VB.Menu mnuPlanName 
      Caption         =   "方案名称"
      Visible         =   0   'False
      Begin VB.Menu mnuPlanNameNew 
         Caption         =   "新增方案(&N)"
      End
      Begin VB.Menu mnuPlanNameUpdate 
         Caption         =   "修改方案(&U)"
      End
      Begin VB.Menu mnuPlanNameRemove 
         Caption         =   "删除方案(&R)"
      End
      Begin VB.Menu mnuPlanNameSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlanNameStart 
         Caption         =   "启用方案(&S)"
      End
      Begin VB.Menu mnuPlanNameStop 
         Caption         =   "停用方案(&T)"
      End
   End
   Begin VB.Menu mnuPlanDetail 
      Caption         =   "方案内容"
      Visible         =   0   'False
      Begin VB.Menu mnuPlanDetailAdd 
         Caption         =   "新增时间段(&A)"
      End
      Begin VB.Menu mnuPlanDetailModify 
         Caption         =   "修改时间段(&M)"
      End
      Begin VB.Menu mnuPlanDetailDel 
         Caption         =   "删除时间段(&D)"
      End
   End
End
Attribute VB_Name = "frmRunLimitPlanManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsPlan As ADODB.Recordset
Private mlngPlanNo As Long
Private mobjBar As CommandBar
Private mobjMenu As CommandBarPopup
Private mobjPopup As CommandBarPopup
Private mobjControl As CommandBarControl
Private Const vsfTitleBackColor = &HF0E5BD  '方案内容表格标题背景颜色
Private Const vsfContentBackColor = &HFFFAE4 '方案内容表格内容部分浅色背景色
Private Const HighlightForeColor = &H80000005  '高亮前景色
Private Const HighlightBackColor = &H8000000D  '高亮背景色
Private Const vsfTitleHeight = 500
Private Const vsfRowHeight = 1000
Private Enum PlanDetailTitle
    PDT_星期 = 0
    PDT_时间段1 = 1
    PDT_时间段扩展 = 2
End Enum
Private Enum PlanDetail
    PD_标题 = 0
    PD_星期日 = 1
    PD_星期一 = 2
    PD_星期二 = 3
    PD_星期三 = 4
    PD_星期四 = 5
    PD_星期五 = 6
    PD_星期六 = 7
End Enum

Private Enum CbsMainId
    CMI_Exit = 11
    CMI_NewPlan = 21
    CMI_UpdatePlan = 22
    CMI_RemovePlan = 23
    CMI_StartPlan = 24
    CMI_StopPlan = 25
    CMI_AddTime = 26
    CMI_EditTime = 27
    CMI_DeleteTime = 28
    CMI_ShowStopPlan = 31
End Enum

Public Sub ShowMe(Optional ByVal lngPlanNo As Long)
    '如果有lngPlanNo的话，就选中对应方案
    mlngPlanNo = lngPlanNo
    If mlngPlanNo = 0 Then mlngPlanNo = 1
    Me.Show vbModal, frmMDIMain
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errH
    Select Case Control.id
        Case CMI_Exit
            '退出
            Unload Me
        Case CMI_NewPlan
            '新增方案
            Call NewPlan
        Case CMI_UpdatePlan
            '修改方案
            Call UpdatePlan
        Case CMI_RemovePlan
            '删除方案
            Call RemovePlan
        Case CMI_StartPlan
            '启用方案
            Call StartPlan
        Case CMI_StopPlan
            '停用方案
            Call StopPlan
        Case CMI_AddTime
            '新增时间段
            Call AddTime
        Case CMI_EditTime
            '修改时间段
            Call EditTime
        Case CMI_DeleteTime
            '删除时间段
            Call DeleteTime
        Case CMI_ShowStopPlan
            '显示停用方案
            Control.Checked = Not Control.Checked
            Call ShowStopPlan
    End Select
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case CMI_UpdatePlan
            '修改方案
            Control.Enabled = mnuPlanNameUpdate.Enabled
        Case CMI_RemovePlan
            '删除方案
            Control.Enabled = mnuPlanNameRemove.Enabled
        Case CMI_StartPlan
            '启用方案
            Control.Enabled = mnuPlanNameStart.Enabled
        Case CMI_StopPlan
            '停用方案
            Control.Enabled = mnuPlanNameStop.Enabled
        Case CMI_AddTime
            '新增时间段
            Control.Enabled = mnuPlanDetailAdd.Enabled
        Case CMI_EditTime
            '修改时间段
            Control.Enabled = mnuPlanDetailModify.Enabled
        Case CMI_DeleteTime
            '删除时间段
            Control.Enabled = mnuPlanDetailDel.Enabled
    End Select
End Sub

Private Sub Form_Load()
    Call InitCbsMain
    Call FillPlanList
    Call FormatVsfPlan
End Sub

Private Sub InitCbsMain()
    With CommandBarsGlobalSettings
        Set .App = App
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll"                                       '设置中文语言资源文件
        .ColorManager.SystemTheme = xtpSystemThemeAuto                                          '控件整体的颜色方案
    End With
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False                                                         '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True                                                                '显示按钮提示
        .AlwaysShowFullMenus = False                                                            '不常用的菜单项先隐藏
        .UseFadedIcons = True                                                                   '图标显示为褪色效果
        .IconsWithShadow = True                                                                 '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True                                                                '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True                                                                      '工具栏显示为大图标
        .SetIconSize True, 24, 24                                                               '设置大图标的尺寸
        .SetIconSize False, 16, 16                                                              '设置小图标的尺寸
    End With
    
    With cbsMain
        .VisualTheme = xtpThemeOffice2003                                                       '设置控件显示风格
        .EnableCustomization False                                                              '是否允许自定义设置
        Set cbsMain.Icons = ImageMan.Icons                                                      '设置关联的图标控件
    End With                                                                                    '菜单宽度自动拉申且宽度不足时也不换行
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 1, "文件(&F)", -1, False) '菜单栏定义
    With mobjMenu.CommandBar.Controls
        Set mobjControl = .Add(xtpControlButton, 11, "退出(&X)")
    End With

    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 2, "编辑(&E)", -1, False)
    With mobjMenu.CommandBar.Controls
        Set mobjControl = .Add(xtpControlButton, 21, "新增方案(&N)")
        Set mobjControl = .Add(xtpControlButton, 22, "修改方案(&U)")
        Set mobjControl = .Add(xtpControlButton, 23, "删除方案(&R)")
        
        Set mobjControl = .Add(xtpControlButton, 24, "启用方案(&S)")
        mobjControl.BeginGroup = True
        Set mobjControl = .Add(xtpControlButton, 25, "停用方案(&T)")
        
        Set mobjControl = .Add(xtpControlButton, 26, "新增时间段(&A)")
        mobjControl.BeginGroup = True
        Set mobjControl = .Add(xtpControlButton, 27, "修改时间段(&M)")
        Set mobjControl = .Add(xtpControlButton, 28, "删除时间段(&D)")
    End With
    
    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 3, "查看(&V)", -1, False)
    With mobjMenu.CommandBar.Controls
        Set mobjControl = .Add(xtpControlButton, 31, "显示停用方案(&S)")
        mobjControl.Checked = False
    End With
'
    Set mobjBar = cbsMain.Add("工具栏", xtpBarTop)                                               '工具栏定义
    With mobjBar.Controls
        Set mobjControl = .Add(xtpControlButton, 21, "新增方案")
        mobjControl.Style = xtpButtonIconAndCaption
        Set mobjControl = .Add(xtpControlButton, 22, "修改方案")
        mobjControl.Style = xtpButtonIconAndCaption
        Set mobjControl = .Add(xtpControlButton, 23, "删除方案")
        mobjControl.Style = xtpButtonIconAndCaption
        
        Set mobjControl = .Add(xtpControlButton, 24, "启用方案")
        mobjControl.Style = xtpButtonIconAndCaption
        mobjControl.BeginGroup = True
        Set mobjControl = .Add(xtpControlButton, 25, "停用方案")
        mobjControl.Style = xtpButtonIconAndCaption
        
        Set mobjControl = .Add(xtpControlButton, 26, "新增时间段")
        mobjControl.Style = xtpButtonIconAndCaption
        mobjControl.BeginGroup = True
        Set mobjControl = .Add(xtpControlButton, 27, "修改时间段")
        mobjControl.Style = xtpButtonIconAndCaption
        Set mobjControl = .Add(xtpControlButton, 28, "删除时间段")
        mobjControl.Style = xtpButtonIconAndCaption
        
        Set mobjControl = .Add(xtpControlButton, 11, "退出")
        mobjControl.Style = xtpButtonIconAndCaption
        mobjControl.BeginGroup = True
    End With
End Sub

Private Sub FillPlanList()
    '填充左下方方案列表
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim objNode As Node
    Dim i As Long
    
    On Error GoTo errH
    strSql = "Select 序号, 名称, 是否启用, 描述 From ZlRunLimit Order by 是否启用 Desc, 序号"
    Set mrsPlan = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption)
    tvwPlanTree.Nodes.Clear
    With mrsPlan
        Do While Not .EOF
            If !是否启用 = 1 Then
                If !名称 = "预设方案" Then
                    Set objNode = tvwPlanTree.Nodes.Add(, , "K_" & !序号, !名称, "enabledLock")
                Else
                    Set objNode = tvwPlanTree.Nodes.Add(, , "K_" & !序号, !名称, "enabled")
                End If
                objNode.Tag = !描述 & ""
            ElseIf mnuViewShow.Checked Then
                If !名称 = "预设方案" Then
                    Set objNode = tvwPlanTree.Nodes.Add(, , "K_" & !序号, !名称, "disabledLock")
                Else
                    Set objNode = tvwPlanTree.Nodes.Add(, , "K_" & !序号, !名称, "disabled")
                End If
                objNode.Tag = !描述 & ""
            End If
            .MoveNext
        Loop
        On Error Resume Next
            tvwPlanTree.Nodes("K_" & mlngPlanNo).Selected = True
            Call tvwPlanTree_NodeClick(tvwPlanTree.Nodes("K_" & mlngPlanNo))
        If err.Number <> 0 Then
            If tvwPlanTree.Nodes.Count > 0 Then
                mlngPlanNo = Split(tvwPlanTree.Nodes(1).Key, "_")(1)
                tvwPlanTree.Nodes("K_" & mlngPlanNo).Selected = True
                Call tvwPlanTree_NodeClick(tvwPlanTree.Nodes("K_" & mlngPlanNo))
            Else
                Call ClearPlanDetail
                Call SetEnabled(False)
            End If
            err.Clear
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FormatVsfPlan()
    '设置右侧方案展示表格格式
    With vsfPlanDetail
        .Cell(flexcpPicture, 0, 0) = imgPlanDetail.ListImages(1).Picture
        .GridLines = flexGridNone
        .rowHeight(PD_标题) = vsfTitleHeight
        .rowHeight(PD_星期日) = vsfRowHeight
        .rowHeight(PD_星期一) = vsfRowHeight
        .rowHeight(PD_星期二) = vsfRowHeight
        .rowHeight(PD_星期三) = vsfRowHeight
        .rowHeight(PD_星期四) = vsfRowHeight
        .rowHeight(PD_星期五) = vsfRowHeight
        .rowHeight(PD_星期六) = vsfRowHeight
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tvwPlanTree.Height = Me.ScaleHeight - tvwPlanTree.Top - 50
    vsfPlanDetail.Height = Me.ScaleHeight - vsfPlanDetail.Top - 50
    vsfPlanDetail.Left = tvwPlanTree.Left + tvwPlanTree.Width + 50
    vsfPlanDetail.Width = Me.ScaleWidth - vsfPlanDetail.Left - 50
    Call AdjustFormDisplay
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub AdjustFormDisplay()
    With vsfPlanDetail
        .Select 0, 0, .Rows - 1, .Cols - 1
        .CellBorder &HE9D2A5, 1, 0, 1, 2, 2, 2
        .Cell(flexcpBackColor, PD_标题, PDT_星期, 0, .Cols - 1) = vsfTitleBackColor
        .Cell(flexcpBackColor, PD_标题, PDT_星期, .Rows - 1, 0) = vsfTitleBackColor
        .Cell(flexcpBackColor, PD_星期日, PDT_时间段1, PD_星期日, .Cols - 1) = vsfContentBackColor
        .Cell(flexcpBackColor, PD_星期二, PDT_时间段1, PD_星期二, .Cols - 1) = vsfContentBackColor
        .Cell(flexcpBackColor, PD_星期四, PDT_时间段1, PD_星期四, .Cols - 1) = vsfContentBackColor
        .Cell(flexcpBackColor, PD_星期六, PDT_时间段1, PD_星期六, .Cols - 1) = vsfContentBackColor
    End With
End Sub

Private Sub AddTime()
'新增时间段
    Dim strTimeStart As String, strTimeStop As String
    Dim lngRow As Long, lngCol As Long
    Dim j As Long
    
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Text = "预设方案" Then Exit Sub
    With vsfPlanDetail
        If .Row = 0 Or .Row = 8 Then Exit Sub
        lngRow = .Row
        lngCol = .Col
        .Cell(flexcpBackColor, lngRow, lngCol) = HighlightBackColor
        .Cell(flexcpForeColor, lngRow, lngCol) = HighlightForeColor
        If frmRunLimitTimeEdit.ShowMe(0, Mid(tvwPlanTree.SelectedItem.Key, 3), lngRow, strTimeStart, strTimeStop) Then
            Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
            If lngRow > .Rows - 1 Then lngRow = .Rows - 1
            If lngCol > .Cols - 1 Then lngCol = .Cols - 1
            .Row = lngRow
            .Col = lngCol
        End If
        .Cell(flexcpBackColor, lngRow, lngCol) = &HFFF6D2
        .Cell(flexcpForeColor, lngRow, lngCol) = &H80000008
    End With
End Sub

Private Sub DeleteTime()
'删除时间段
    Dim strTimeStart As String, strTimeStop As String
    Dim lngRow As Long, lngCol As Long
    
    On Error GoTo errH
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Text = "预设方案" Then Exit Sub
    With vsfPlanDetail
        If .Row = 0 Or .Row = 8 Or .Col = 0 Or .TextMatrix(.Row, .Col) = "" Then Exit Sub
        strTimeStart = Mid(Split(.TextMatrix(.Row, .Col), vbNewLine)(0), 3)
        strTimeStop = Mid(Split(.TextMatrix(.Row, .Col), vbNewLine)(2), 3)
        lngRow = .Row
        lngCol = .Col
        If MsgBox("确定要将时间段“" & strTimeStart & "-" & strTimeStop & "”删除吗？", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            Call ExecuteProcedure("Zl_ZlRunLimitTime_Update(2," & .Cell(flexcpData, .Row, .Col) & ")", "删除时间段")
            Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
            If lngRow > .Rows - 1 Then lngRow = .Rows - 1
            If lngCol > .Cols - 1 Then lngCol = .Cols - 1
            .Row = lngRow
            .Col = lngCol
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub EditTime()
'修改时间段
    Dim strTimeStart As String, strTimeStop As String
    Dim lngRow As Long, lngCol As Long
    
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Text = "预设方案" Then Exit Sub
    With vsfPlanDetail
        If .Row = 0 Or .Col = 0 Or .TextMatrix(.Row, .Col) = "" Then Exit Sub
        strTimeStart = Mid(Split(.TextMatrix(.Row, .Col), vbNewLine)(0), 3)
        strTimeStop = Mid(Split(.TextMatrix(.Row, .Col), vbNewLine)(2), 3)
        lngRow = .Row
        lngCol = .Col
        .Cell(flexcpBackColor, lngRow, lngCol) = HighlightBackColor
        .Cell(flexcpForeColor, lngRow, lngCol) = HighlightForeColor
        If frmRunLimitTimeEdit.ShowMe(.Cell(flexcpData, .Row, .Col), Mid(tvwPlanTree.SelectedItem.Key, 3), .Row, strTimeStart, strTimeStop) Then
            Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
            If lngRow > .Rows - 1 Then lngRow = .Rows - 1
            If lngCol > .Cols - 1 Then lngCol = .Cols - 1
            .Row = lngRow
            .Col = lngCol
        End If
        .Cell(flexcpBackColor, lngRow, lngCol) = &HFFF6D2
        .Cell(flexcpForeColor, lngRow, lngCol) = &H80000008
    End With
End Sub

Private Sub NewPlan()
'新增方案
    Dim lngPlanNo As Long
    Dim objNode As Node
    Dim strPlanName As String, strDescription As String
    
    If frmRunLimitPlanEdit.ShowMe(Me, lngPlanNo, strPlanName, strDescription) Then
        Set objNode = tvwPlanTree.Nodes.Add(, , "K_" & lngPlanNo, strPlanName, "enabled")
        objNode.Tag = strDescription
        tvwPlanTree.Nodes("K_" & lngPlanNo).Selected = True
        Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
    End If
End Sub

Private Sub RemovePlan()
'删除方案
    On Error GoTo errH
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    '判断是否有模块正在使用该方案，并弹出提示
    If CheckPlanStatus("删除") = False Then Exit Sub
    
    If MsgBox("确定要删除该方案吗？", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) Then
        Call ExecuteProcedure("Zl_Zlrunlimit_Update(2," & mlngPlanNo & ")", "删除方案")
        tvwPlanTree.Nodes.Remove (tvwPlanTree.SelectedItem.Key)
        If tvwPlanTree.Nodes.Count > 0 Then
            tvwPlanTree.Tag = ""
            Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
        Else
            Call SetEnabled(False)
            Call ClearPlanDetail
        End If
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub StartPlan()
'启用方案
    On Error GoTo errH
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Image = "enabled" Or tvwPlanTree.SelectedItem.Image = "enabledLock" Then Exit Sub
    Call ExecuteProcedure("Zl_Zlrunlimit_Update(1," & mlngPlanNo & ",Null,1)", "启用方案")
    If tvwPlanTree.SelectedItem.Image = "disabled" Then
        tvwPlanTree.SelectedItem.Image = "enabled"
    Else
        tvwPlanTree.SelectedItem.Image = "enabledLock"
    End If
    Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub StopPlan()
'停用方案
    On Error GoTo errH
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Image = "disabled" Or tvwPlanTree.SelectedItem.Image = "disabledLock" Then Exit Sub
    '判断是否有模块正在使用该方案，并弹出提示
    If CheckPlanStatus("停用") = False Then Exit Sub
    Call ExecuteProcedure("Zl_Zlrunlimit_Update(1," & mlngPlanNo & ",Null,0)", "停用方案")
    If tvwPlanTree.SelectedItem.Image = "enabled" Then
        tvwPlanTree.SelectedItem.Image = "disabled"
    Else
        tvwPlanTree.SelectedItem.Image = "disabledLock"
    End If
    If mnuViewShow.Checked = False Then
        tvwPlanTree.Nodes.Remove (tvwPlanTree.SelectedItem.Key)
        If tvwPlanTree.Nodes.Count > 0 Then
            tvwPlanTree.Tag = ""
            Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
        Else
            Call SetEnabled(False)
            Call ClearPlanDetail
        End If
    Else
        Call tvwPlanTree_NodeClick(tvwPlanTree.SelectedItem)
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'检查方案是否正在被使用
Private Function CheckPlanStatus(ByVal strTag As String) As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strFuncList As String
    Dim i As Long
    
    On Error GoTo errH
    strSql = "Select a.功能 From Zlrunlimitset A, Zlrunlimit B Where a.方案序号 = b.序号 And b.序号 =[1]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption, mlngPlanNo)
    With rsTemp
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
                If i > 3 Then Exit For
                strFuncList = strFuncList & "“" & !功能 & "”" & vbNewLine
                .MoveNext
            Next
            If .RecordCount > 3 Then
                MsgBox "该限时方案正在被" & vbNewLine & strFuncList & "等" & .RecordCount & _
                "个功能使用，要" & strTag & "该方案请先修改以上功能的方案！", vbInformation, gstrSysName
            Else
                MsgBox "该限时方案正在被功能" & vbNewLine & strFuncList & _
                "使用，要" & strTag & "该方案请先修改以上功能的方案！", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    End With
    CheckPlanStatus = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub UpdatePlan()
'修改方案
    Dim strPlanName As String, strDescription As String
    
    If tvwPlanTree.SelectedItem Is Nothing Then Exit Sub
    If tvwPlanTree.SelectedItem.Text = "预设方案" Then Exit Sub
    strPlanName = tvwPlanTree.SelectedItem.Text
    strDescription = tvwPlanTree.SelectedItem.Tag
    If frmRunLimitPlanEdit.ShowMe(Me, mlngPlanNo, strPlanName, strDescription) Then
        tvwPlanTree.Nodes("K_" & mlngPlanNo).Text = strPlanName
        tvwPlanTree.Nodes("K_" & mlngPlanNo).Tag = strDescription
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuPlanDetailAdd_Click()
    Call AddTime
End Sub

Private Sub mnuPlanDetailDel_Click()
    Call DeleteTime
End Sub

Private Sub mnuPlanDetailModify_Click()
    Call EditTime
End Sub

Private Sub mnuPlanNameNew_Click()
    Call NewPlan
End Sub

Private Sub mnuPlanNameRemove_Click()
    Call RemovePlan
End Sub

Private Sub mnuPlanNameStart_Click()
    Call StartPlan
End Sub

Private Sub mnuPlanNameStop_Click()
    Call StopPlan
End Sub

Private Sub mnuPlanNameUpdate_Click()
    Call UpdatePlan
End Sub

Private Sub ShowStopPlan()
    mnuViewShow.Checked = Not mnuViewShow.Checked
    If tvwPlanTree.Nodes.Count > 0 Then
        mlngPlanNo = Split(tvwPlanTree.Nodes(1).Key, "_")(1)
    End If
    Call FillPlanList
End Sub

Private Sub tvwPlanTree_DblClick()
    Call UpdatePlan
End Sub

Private Sub tvwPlanTree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPlanName
    End If
End Sub

Private Sub tvwPlanTree_NodeClick(ByVal Node As MSComctlLib.Node)
    If tvwPlanTree.Tag <> "" Then
        tvwPlanTree.Nodes(tvwPlanTree.Tag).BackColor = &H80000005
        tvwPlanTree.Nodes(tvwPlanTree.Tag).ForeColor = &H80000012
    End If
    Node.BackColor = HighlightBackColor
    Node.ForeColor = HighlightForeColor
    tvwPlanTree.Tag = tvwPlanTree.SelectedItem.Key
    mlngPlanNo = Split(Node.Key, "_")(1)
    If Node.Text = "预设方案" Then
        Call SetEnabled(False)
    Else
        Call SetEnabled(True)
    End If
    Call FillPlanDetail
End Sub

Private Sub SetEnabled(ByVal blnEnabled As Boolean)
    mnuPlanNameUpdate.Enabled = blnEnabled
    mnuPlanNameRemove.Enabled = blnEnabled
    mnuPlanDetailAdd.Enabled = blnEnabled
    mnuPlanDetailModify.Enabled = blnEnabled
    mnuPlanDetailDel.Enabled = blnEnabled
    If tvwPlanTree.Nodes.Count = 0 Then
        mnuPlanNameStart.Enabled = False
        mnuPlanNameStop.Enabled = False
    Else
        If tvwPlanTree.SelectedItem.Image = "enabledLock" Or tvwPlanTree.SelectedItem.Image = "enabled" Then
            mnuPlanNameStart.Enabled = False
            mnuPlanNameStop.Enabled = True
        Else
            mnuPlanNameStart.Enabled = True
            mnuPlanNameStop.Enabled = False
        End If
    End If
End Sub

Private Sub FillPlanDetail()
'填充详细方案信息
    Dim j As Long  '表示时间段
    Dim lngLastWeekNo As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
        '将老的方案信息清空
        Call ClearPlanDetail
        
        '填充新方案
        strSql = "Select Id, 星期, To_Char(开始时间, 'HH24:MI:SS') 开始时间, To_Char(结束时间, 'HH24:MI:SS') 结束时间" & vbNewLine & _
                "From ZlRunLimitTime" & vbNewLine & _
                "Where 方案 = [1]" & vbNewLine & _
                "Order By 星期, 开始时间, 结束时间"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption, mlngPlanNo)
        With rsTemp
            Do While Not .EOF
                If !星期 = lngLastWeekNo Then
                    j = j + 1
                    If j + 2 > vsfPlanDetail.Cols Then
                        vsfPlanDetail.Cols = j + 2
                        vsfPlanDetail.ColWidth(j) = vsfPlanDetail.ColWidth(PDT_时间段1)
                        vsfPlanDetail.TextMatrix(0, j) = "时间段" & j
                        vsfPlanDetail.ColAlignment(j) = flexAlignCenterCenter
                        Call AdjustFormDisplay
                    End If
                Else
                    j = 1
                End If
                vsfPlanDetail.TextMatrix(!星期 + 1, j) = "起 " & !开始时间 & vbNewLine & vbNewLine & "止 " & !结束时间
                vsfPlanDetail.Cell(flexcpData, !星期 + 1, j) = Val(!id & "")
                lngLastWeekNo = !星期
                .MoveNext
            Loop
        End With
        vsfPlanDetail.ToolTipText = tvwPlanTree.SelectedItem.Tag
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'将老的方案信息清空
Private Sub ClearPlanDetail()
    Dim i As Long
    
    With vsfPlanDetail
        .Cols = 3
        .TextMatrix(0, PDT_时间段扩展) = ""
        For i = 1 To 7
            .TextMatrix(i, PDT_时间段1) = ""
            .TextMatrix(i, PDT_时间段扩展) = ""
        Next
        Call AdjustFormDisplay
    End With
End Sub

Private Sub vsfPlanDetail_DblClick()
    With vsfPlanDetail
        If .MouseRow <> .Row Then Exit Sub
        If .TextMatrix(.Row, .Col) = "" Then
            Call AddTime
        Else
            Call EditTime
        End If
    End With
End Sub

Private Sub vsfPlanDetail_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        Call DeleteTime
    End Select
End Sub

Private Sub vsfPlanDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPlanDetail
    End If
End Sub
