VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExaminePathTable 
   AutoRedraw      =   -1  'True
   Caption         =   "临床路径审核"
   ClientHeight    =   9495
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   14550
   Icon            =   "frmExaminePathTable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14550
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   6000
      ScaleHeight     =   2385
      ScaleWidth      =   4680
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4680
      Begin VSFlex8Ctl.VSFlexGrid vsgIllness 
         Height          =   855
         Left            =   480
         TabIndex        =   9
         Top             =   1440
         Width           =   5535
         _cx             =   9763
         _cy             =   1508
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   1
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
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "●适用科室："
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   13
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "………………………………………………………………………………"
         Height          =   180
         Index           =   1
         Left            =   330
         MouseIcon       =   "frmExaminePathTable.frx":058A
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   780
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl病种 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "●对应病种："
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   11
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lbl说明 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明：…………………………………………………………………………………"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   6210
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   360
      ScaleHeight     =   6015
      ScaleWidth      =   4215
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
      Begin VB.PictureBox picDetail 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1245
         Left            =   600
         ScaleHeight     =   1245
         ScaleWidth      =   2895
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3840
         Width           =   2895
         Begin XtremeReportControl.ReportControl rptLog 
            Height          =   1095
            Left            =   960
            TabIndex        =   7
            Top             =   480
            Width           =   3015
            _Version        =   589884
            _ExtentX        =   5318
            _ExtentY        =   1931
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
      End
      Begin VB.PictureBox picList 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   2085
         Left            =   120
         ScaleHeight     =   2085
         ScaleWidth      =   3615
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   3615
         Begin XtremeReportControl.ReportControl rptPath 
            Height          =   1095
            Left            =   -120
            TabIndex        =   5
            Top             =   120
            Width           =   3015
            _Version        =   589884
            _ExtentX        =   5318
            _ExtentY        =   1931
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
      End
      Begin XtremeSuiteControls.TabControl tbcPath 
         Height          =   675
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   1935
         _Version        =   589884
         _ExtentX        =   3413
         _ExtentY        =   1191
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2250
      ScaleHeight     =   600
      ScaleWidth      =   660
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   660
   End
   Begin MSComctlLib.ImageList ilsPic 
      Left            =   1245
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExaminePathTable.frx":06DC
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExaminePathTable.frx":0C76
            Key             =   "File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExaminePathTable.frx":1210
            Key             =   "branch"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExaminePathTable.frx":7A72
            Key             =   "Merge"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   9135
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmExaminePathTable.frx":E2D4
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22754
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcContent 
      Height          =   675
      Left            =   6840
      TabIndex        =   14
      Top             =   6360
      Width           =   1935
      _Version        =   589884
      _ExtentX        =   3413
      _ExtentY        =   1191
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   480
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmExaminePathTable.frx":EB66
      Left            =   1320
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmExaminePathTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mfrmDesign As frmPathDesign
Attribute mfrmDesign.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmPathDesign
Attribute mfrmContent.VB_VarHelpID = -1
Private WithEvents mfrmEdit As frmPathEdit
Attribute mfrmEdit.VB_VarHelpID = -1
Private mstrPrivs As String
Private mlngModul As Long
Private mlng路径ID As Long
Private mlng版本号 As Long
Private mint审核状态 As Integer  '审核状态：0-编辑;1-提交审核;2-药剂科审核通过;3-药剂科拒绝通过；4-医务科通过；5-医务科拒绝通过

Private Enum E_STATUS
    E_编辑 = 0
    E_提交 = 1
    E_药剂通过 = 2
    E_药剂拒绝 = 3
    E_通过 = 4
    E_拒绝 = 5
End Enum

Private Enum COL_LIST
    COL_ID = 0
    COL_图标 = 1
    COL_分支 = 2
    COL_行号 = 3
    COL_分类 = 4
    COL_编码 = 5
    COL_名称 = 6
    COL_审核状态
    COL_病例分型
    COL_适用病情
    COL_适用性别
    COL_适用年龄
    COL_说明
    COL_通用
    COL_最新版本
    COL_审核类型
    COL_性质     '1=合并路径 0=首要路径
    COL_版本号
End Enum

Private Enum COL_LIST_LOG
    LOG_类型 = 0
    LOG_操作说明
    LOG_操作人员
    LOG_操作时间
End Enum


Private Enum CHK_INDEX
    CHK_已经审核 = 0
    CHK_未审核 = 1
    CHK_已停用 = 2
End Enum

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim lng路径ID As Long
    Dim blnTmp As Boolean
    Dim str分类 As String
    Dim str编码 As String
    Dim frmSub As Form
    
    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    Select Case Control.ID
    Case conMenu_Edit_Audit     '审核\医务科审核
        Call FuncVersionAudit(1)
    Case conMenu_Edit_Untread '取消 审核\医务科审核
        Call FuncVersionAudit(2)
    Case conMenu_Edit_Preferences '全路径项目
        Call frmPathItemAll.ShowMe(Me, mstrPrivs, mlng路径ID, mlng版本号, True)
    Case conMenu_View_Refresh    '刷新
        Call RefreshData
    Case conMenu_Edit_Modify  '查看目录
        Call FuncPathView
    Case conMenu_File_Exit    '退出
        Unload Me
    End Select
End Sub

Public Sub ShowMe(frmParent As Object, ByVal lngMode As Long, ByVal strPrivs As String)
    mstrPrivs = strPrivs
    mlngModul = lngMode
    gbln双审核 = zlDatabase.GetPara("双审核模式", glngSys, p临床路径管理) = 1
    '非双审核模式
        gbln双审核 = False
        mstrPrivs = ";基本;全院路径;审核;"
    '非审核模式
'        gbln双审核 = True
'        mstrPrivs = ";基本;全院路径;审核;药剂科审核;"
    Me.Show 1, frmParent
End Sub

Private Sub FuncPathView()
    mfrmEdit.ShowEdit Me, mstrPrivs, mlng路径ID, , True
End Sub

Private Sub FuncVersionAudit(ByVal bytFunc As Byte)
'功能：审核/取消审核当前版本
'参数
'   bytFunc 1=审核;2=取消审核
'   1=医务科审核 -1=医务科取消审核 2=药剂科审核 -2=药剂科取消审核

    Dim strSql As String
    Dim intNum As Integer
    If bytFunc = 1 Then
        If gbln双审核 Then
            If mint审核状态 = E_提交 Then
                intNum = 2
            ElseIf mint审核状态 = E_药剂通过 Then
                intNum = 1
            End If
        Else
            intNum = 1
        End If
    ElseIf bytFunc = 2 Then
        If gbln双审核 Then
            If mint审核状态 = E_通过 Or mint审核状态 = E_拒绝 Then
                intNum = 5
            ElseIf mint审核状态 = E_药剂通过 Or mint审核状态 = E_药剂拒绝 Then
                intNum = 6
            End If
        Else
            intNum = 5
        End If
    End If
    If intNum = 5 Or intNum = 6 Then
        If MsgBox("确实要取消审核当前版本的临床路径吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        strSql = "Zl_临床路径审核_Insert(" & intNum & "," & mlng路径ID & "," & mlng版本号 & ",NULL,NULL," & IIf(gbln双审核, 1, 0) & ")"
        On Error GoTo errH
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        On Error GoTo 0
    ElseIf intNum = 1 Or intNum = 2 Then
        If frmPathAduit.ShowAudit(Me, mlng路径ID, mlng版本号, intNum) = False Then Exit Sub
    End If
    
    Call RefreshData
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.staThis.Visible Then Bottom = Me.staThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    picLeft.Move lngLeft, lngTop, , lngBottom - lngTop
    picDetail.Move lngLeft, lngTop, , lngBottom - lngTop
    With Me.picInfo
        .Left = picLeft.Left + picLeft.Width + 60
        .Top = lngTop
        .Width = lngRight - .Left
        .Height = lngBottom - lngTop
    End With
    Call ResizeInfoPane
    With Me.tbcContent
        .Left = picInfo.Left
        .Top = picInfo.Top + picInfo.Height
        .Width = picInfo.Width
        .Height = lngBottom - .Top
    End With
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean

    Select Case Control.ID
    Case conMenu_Edit_Audit
        If (InStr(";" & mstrPrivs & ";", ";审核;") = 0 And InStr(";" & mstrPrivs & ";", ";药剂科审核;") = 0 And gbln双审核) Or _
            (gbln双审核 = False And InStr(";" & mstrPrivs & ";", ";审核;") = 0) Then
            Control.Visible = False
        Else
            blnEnabled = mlng路径ID <> 0 And mlng版本号 > 0 And (mint审核状态 = E_药剂通过 Or mint审核状态 = E_提交) And tbcPath.Selected.Tag = "待审核"
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_Untread
        If (InStr(";" & mstrPrivs & ";", ";审核;") = 0 And InStr(";" & mstrPrivs & ";", ";药剂科审核;") = 0 And gbln双审核) Or _
            (gbln双审核 = False And InStr(";" & mstrPrivs & ";", ";审核;") = 0) Then
            Control.Visible = False
        Else
            blnEnabled = mlng路径ID <> 0 And mlng版本号 > 0 And InStr(",2,3,4,5,", mint审核状态) > 0 And InStr(",审核通过,审核未过,", tbcPath.Selected.Tag) > 0
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_Modify
        Control.Enabled = mlng路径ID <> 0 And mlng版本号 > 0
    Case conMenu_Edit_Preferences '全路径项目
        Control.Enabled = mlng路径ID <> 0 And mlng版本号 > 0
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picLeft.Hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picDetail.Hwnd
    End If
End Sub

Private Sub Form_Load()
    Dim objPane As XtremeDockingPane.Pane
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    gbln双审核 = zlDatabase.GetPara("双审核模式", glngSys, p临床路径管理) = 1
    Call zlCommFun.SetWindowsInTaskBar(Me.Hwnd, False)

    Set mfrmEdit = New frmPathEdit
    Set mfrmDesign = New frmPathDesign
    Set mfrmContent = New frmPathDesign

    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    '审查详情
    '---------------------------------------------------
    Call InitReportColumnLog
    
    Call MainDefCommandBar
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 400, 300, DockLeftOf, Nothing)
    objPane.Title = "路径列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set objPane = Me.dkpMain.CreatePane(2, 400, 200, DockBottomOf, objPane)
    objPane.Title = "审核列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
     
    'tbcPath 路径列表
    With Me.tbcPath
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "待审核", picList.Hwnd, 0).Tag = "待审核"
        .InsertItem(1, "审核通过", picList.Hwnd, 0).Tag = "审核通过"
        .InsertItem(2, "审核未过", picList.Hwnd, 0).Tag = "审核未过"
        
        .Item(2).Selected = True
        .Item(0).Selected = True
        '定位路径选项卡
        .Item(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPath), "tbcPath", 0)).Selected = True
    End With
        
    'TabControl
    '-----------------------------------------------------
    With Me.tbcContent
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
        End With
        .InsertItem 0, "临床路径表", mfrmContent.Hwnd, 0
    End With
 
    '对应病种
    '---------------------------------------------------------
    Call InitVsgIllness
    
    Call RestoreWinState(Me, App.ProductName)
    Call RefreshData
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    Dim lngCount As Long
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True    '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审核"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消审核")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "路径信息")
        objControl.IconId = 3022
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Preferences, "全路径项目")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        objControl.BeginGroup = True
    End With
    
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh    '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help    '帮助
    End With

End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptPath
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)或ItemIndex查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(COL_ID, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_图标, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_分支, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_行号, "行号", 35, True)
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_分类, "分类", 80, True)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_编码, "编码", 50, True)
        objCol.Groupable = False
        Set objCol = .Columns.Add(COL_名称, "名称", 150, True)
        objCol.Groupable = False
        Set objCol = .Columns.Add(COL_审核状态, "审核进度", 120, True)
        Set objCol = .Columns.Add(COL_病例分型, "病例分型", 65, True)
        Set objCol = .Columns.Add(COL_适用病情, "适用病情", 55, True)
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_适用性别, "适用性别", 55, True)
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_适用年龄, "适用年龄", 55, True)
        Set objCol = .Columns.Add(COL_说明, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_通用, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_最新版本, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_审核类型, "审核类型", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_性质, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_版本号, "", 0, False)
        objCol.Visible = False

        For Each objCol In .Columns
            objCol.Editable = False
        Next

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的临床路径..."
            '.ShadeGroupHeadings = True
        End With

        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False    '会引发SelectionChanged事件
        .SetImageList Me.ilsPic
                
        If gbln双审核 And InStr(mstrPrivs, ";药剂科审核;") > 0 And InStr(mstrPrivs, ";审核;") > 0 Then
            .GroupsOrder.Add .Columns(COL_审核类型)
            .GroupsOrder(0).SortAscending = True    '分组之后,如果分组列不显示,分组列的排序是不变的
            .GroupsOrder.Add .Columns(COL_分类)
            .GroupsOrder(0).SortAscending = True
            '分组之后可能失去记录集中的顺序,因此强行加入排序列
            .SortOrder.Add .Columns(COL_审核类型)
            .SortOrder(0).SortAscending = True
            .SortOrder.Add .Columns(COL_分类)
            .SortOrder(1).SortAscending = True
            .SortOrder.Add .Columns(COL_编码)
            .SortOrder(2).SortAscending = True
        Else
            .GroupsOrder.Add .Columns(COL_分类)
            .GroupsOrder(0).SortAscending = True    '分组之后,如果分组列不显示,分组列的排序是不变的
    
            '分组之后可能失去记录集中的顺序,因此强行加入排序列
            .SortOrder.Add .Columns(COL_分类)
            .SortOrder(0).SortAscending = True
            .SortOrder.Add .Columns(COL_编码)
            .SortOrder(1).SortAscending = True
        End If
    End With
End Sub

Private Sub InitReportColumnLog()
    Dim objCol As ReportColumn

    With rptLog
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)或ItemIndex查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(LOG_类型, "审核结果", 100, True)
        Set objCol = .Columns.Add(LOG_操作说明, "审核说明", 200, True)
        Set objCol = .Columns.Add(LOG_操作人员, "审核人", 80, True)
        Set objCol = .Columns.Add(LOG_操作时间, "审核时间", 140, True)

        For Each objCol In .Columns
            objCol.Editable = False
        Next

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的审核内容..."
        End With

        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False    '会引发SelectionChanged事件
        .SetImageList Me.ilsPic
    
        .SortOrder.Add .Columns(LOG_操作时间)
        .SortOrder(0).SortAscending = False
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Unload mfrmDesign
    Set mfrmDesign = Nothing
    
    Unload mfrmContent
    Set mfrmContent = Nothing
    
    Unload mfrmEdit
    Set mfrmEdit = Nothing
End Sub


Private Sub lbl科室_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 And lbl科室(Index).Caption <> "" Then
        Me.lbl科室(1).Font.Underline = True
        Me.lbl科室(1).ForeColor = RGB(0, 0, 128)
    End If
End Sub

Private Sub mfrmDesign_DataChanged(ByVal 路径ID As Long)
'刷新路径表信息
    Call mfrmContent.zlRefresh(路径ID, mstrPrivs, lbl科室(1).Caption, vsgIllness.Tag)
End Sub

Private Sub mfrmEdit_AfterSave(ByVal 分类 As String, ByVal 编码 As String)
    Call RefreshData(分类, 编码)
End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl科室(1).Font.Underline = False
    Me.lbl科室(1).ForeColor = lbl科室(0).ForeColor
    vsgIllness.FontUnderline = False
    vsgIllness.ForeColor = lbl病种(0).ForeColor
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next

    lbl说明.Width = picInfo.ScaleWidth - lbl说明.Left * 2
    lbl科室(1).Width = picInfo.ScaleWidth - lbl科室(1).Left - lbl说明.Left
    vsgIllness.Left = lbl病种(0).Left
    vsgIllness.Width = picInfo.ScaleWidth - vsgIllness.Left - lbl说明.Left
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With tbcPath
        .Top = 0
        .Left = 0
        .Height = picLeft.ScaleHeight
        .Width = picLeft.ScaleWidth
    End With
    picList.Width = picLeft.ScaleWidth
    rptLog.Move 0, 0, picLeft.ScaleWidth
End Sub

Private Sub picDetail_Resize()
    On Error Resume Next
    rptLog.Left = 0
    rptLog.Top = 0
    rptLog.Width = picDetail.ScaleWidth
    rptLog.Height = picDetail.ScaleHeight
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    rptPath.Left = 0
    rptPath.Top = 0
    rptPath.Width = picList.ScaleWidth
    rptPath.Height = picList.ScaleHeight
End Sub

Private Sub rptPath_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow
    
    If KeyCode = vbKeyReturn And Shift = 0 Then
        If rptPath.SelectedRows.count > 0 Then
            If Not rptPath.SelectedRows(0).GroupRow Then
                Set objRow = rptPath.SelectedRows(0)
            End If
        End If
        If Not objRow Is Nothing Then
            Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
            If Not objControl Is Nothing Then objControl.Execute
        End If
    End If
End Sub

Private Sub rptPath_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objControl As CommandBarControl
    
    If Not Row.GroupRow Then
        Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
        If Not objControl Is Nothing Then objControl.Execute
    End If
End Sub

Private Sub rptPath_SelectionChanged()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim arrStr As Variant
    Dim intRowNum As Integer: Dim intColNum As Integer
    Dim i As Long

    On Error GoTo errH

    If rptPath.SelectedRows.count = 0 Then
        Call ClearSubData
    ElseIf rptPath.SelectedRows(0).GroupRow Then
        Call ClearSubData
    Else
        With rptPath.SelectedRows(0)
            mlng路径ID = Val(.Record(COL_ID).Value)
            mlng版本号 = Val(.Record(COL_版本号).Value)
            mint审核状态 = Val(.Record(COL_审核状态).Value)
            
            '对应审查详情
            Call LoadAduit(mlng路径ID, mlng版本号)
            
            lbl说明.Caption = "说明：" & .Record(COL_说明).Value
            
            '对应科室信息
            If .Record(COL_通用).Value = 1 Then
                lbl科室(1).Caption = "该临床路径适用于所有住院临床科室"
            Else
                strSql = "Select B.编码,B.名称 From 临床路径科室 A,部门表 B Where A.科室ID=B.ID And A.路径ID=[1] Order by B.编码"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(.Record(COL_ID).Value))
                strTmp = ""
                Do While Not rsTmp.EOF
                    strTmp = strTmp & "," & rsTmp!编码 & "-" & rsTmp!名称
                    rsTmp.MoveNext
                Loop
                If strTmp <> "" Then
                    lbl科室(1).Caption = Mid(strTmp, 2)
                Else
                    lbl科室(1).Caption = "<该临床路径尚未设置所适用的科室>"
                End If
            End If

            '对应病种信息
            vsgIllness.Clear
        
            strSql = "Select Decode(B.编码,NULL,'['||C.编码||']'||C.名称,'['||B.编码||']'||B.名称) as 名称 ,B.编码 " & _
                     " From 临床路径病种 A,疾病编码目录 B,疾病诊断目录 C" & _
                     " Where A.疾病ID=B.ID(+) And A.诊断ID=C.ID(+) And A.路径ID=[1] and a.性质=0" & _
                     " Order by B.编码,C.编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(.Record(COL_ID).Value))
            strTmp = ""
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!名称
                rsTmp.MoveNext
            Loop
            If strTmp <> "" Then
                With vsgIllness
                    arrStr = Split(Mid(strTmp, 2), ",")
                    .Cols = 3: .Rows = ((UBound(arrStr) + 1) + (.Cols - 1)) \ .Cols
                    .Tag = Mid(strTmp, 2)
                    For i = 0 To UBound(arrStr)
                        intRowNum = i \ .Cols
                        intColNum = i Mod .Cols
                        .TextMatrix(intRowNum, intColNum) = arrStr(i)
                    Next i
                End With
            Else
                vsgIllness.Rows = 1: vsgIllness.Cols = 1
                vsgIllness.TextMatrix(0, 0) = "<该临床路径尚未设置所对应的病种>"
            End If
            
            '路径表信息
            Call mfrmContent.zlRefresh(mlng路径ID, mstrPrivs, lbl科室(1).Caption, vsgIllness.Tag, 2, mlng版本号)
        End With
        
        Call Form_Resize
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ResizeInfoPane()
'功能：根据当前信息内容，调整信息面板内信息项及面板尺寸和位置
'说明：利用Label的AutoSize属性自动调整标签高度
    lbl科室(0).Top = lbl说明.Top + lbl说明.Height + Screen.TwipsPerPixelY * 6
    lbl科室(1).Top = lbl科室(0).Top + lbl科室(0).Height + Screen.TwipsPerPixelY * 3
    lbl病种(0).Top = lbl科室(1).Top + lbl科室(1).Height + Screen.TwipsPerPixelY * 6

    vsgIllness.Top = lbl病种(0).Top + lbl病种(0).Height + Screen.TwipsPerPixelY * 3
    '根据对应病种行数动态显示对应病种信息，最大显示5行
    vsgIllness.Height = vsgIllness.RowHeightMin * IIf(vsgIllness.Rows > 5, 5, vsgIllness.Rows)
    vsgIllness.ColWidthMin = vsgIllness.Width / vsgIllness.Cols
    picInfo.Height = vsgIllness.Top + vsgIllness.Height + Screen.TwipsPerPixelY * 6
End Sub

Private Function RefreshData(Optional ByVal str分类 As String, Optional ByVal str编码 As String) As Boolean
'功能：根据当前设置的条件读取临床路径目录数据
'参数：用于定位
    Dim rsTmp       As ADODB.Recordset
    Dim strSql      As String
    Dim strFilter   As String
    
    Dim objRecord   As ReportRecord
    Dim objItem     As ReportRecordItem
    Dim objRow As ReportRow, i As Long
    Dim lngPreID As Long, lngPreIdx As Long
    Dim intTypeNum  As Integer
    Dim lngPathColor As Long                '未审核路径目录前景颜色值
    Dim lngStopColor As Long                '已停止路径目录前景颜色值
    Dim intType As Integer                 '审核状态
    Dim strTmp As String
    Dim strPrivs As String                  '11-1-药剂科审核;1-医务科审核
    Screen.MousePointer = 11
    
    On Error GoTo errH
    strPrivs = IIf(InStr(mstrPrivs, ";药剂科审核;") > 0, "1", "0") & IIf(InStr(mstrPrivs, ";审核;") > 0, "1", "0")
    
    Select Case tbcPath.Selected.Tag
    Case "待审核"
        If gbln双审核 Then
            If "11" = strPrivs Then
                strFilter = " And Instr(',1,2,',','||a.审核状态||',') > 0"
            ElseIf "10" = strPrivs Then
                strFilter = " And NVL(a.审核状态,0)=1"
            ElseIf "01" = strPrivs Then
                strFilter = " And NVL(a.审核状态,0)=2"
            Else
                strFilter = " And NVL(a.审核状态,0)<0"
            End If
        Else
            '兼容双审核模式切换到非双审核模式,兼容处理审核状态为2的数据
            strFilter = " And Instr(',1,2,',','||a.审核状态||',') > 0"
        End If
    Case "审核通过"
        '兼容历史数据:审核状态为空,审核时间不为空
        If gbln双审核 Then
            If "11" = strPrivs Then
                strFilter = " And (Instr(',2,4,',','||a.审核状态||',') > 0 Or C.审核时间 Is not NULL)  "
            ElseIf "10" = strPrivs Then
                strFilter = " And NVL(a.审核状态,0)=2"
            ElseIf "01" = strPrivs Then
                strFilter = " And (NVL(a.审核状态,0)=4 Or C.审核时间 Is not NULL) "
            Else
                strFilter = " And NVL(a.审核状态,0)<0"
            End If
        Else
            If InStr(mstrPrivs, ";审核;") > 0 Then
                strFilter = " And (NVL(a.审核状态,0)=4 OR C.审核时间 Is Not NULL) "
            Else
                strFilter = " And NVL(a.审核状态,0)<0"
            End If
        End If
    Case "审核未过"
        If gbln双审核 Then
            If "11" = strPrivs Then
                strFilter = " And Instr(',3,5,',','||a.审核状态||',') > 0"
            ElseIf "10" = strPrivs Then
                strFilter = " And NVL(a.审核状态,0)=3"
            ElseIf "01" = strPrivs Then
                strFilter = " And NVL(a.审核状态,0)=5"
            Else
                strFilter = " And NVL(a.审核状态,0)<0"
            End If
        Else
            '兼容双审核模式切换到非双审核模式,兼容处理审核状态为3的数据
            If InStr(mstrPrivs, ";审核;") > 0 Then
                strFilter = " And Instr(',3,5,',','||a.审核状态||',') > 0"
            Else
                strFilter = " And NVL(a.审核状态,0)<0"
            End If
        End If
    End Select

    'SQL中不排序提高效率,ReportControl有排序处理
    strSql = "Select Distinct a.Id, a.分类, a.编码, a.名称, a.病例分型, a.适用病情, a.适用性别, a.适用年龄, a.说明, a.通用, a.最新版本, NVL(a.审核状态,0) as 审核状态," & vbNewLine & _
             "                Decode(b.Id, Null, Null, 1) As 是否存在分支, a.性质,Decode(c.审核时间, Null, 0, 1) As 已审核,c.版本号 " & vbNewLine & _
             "From 临床路径目录 A, 临床路径分支 B, 临床路径版本 C" & vbNewLine & _
             "Where a.Id = b.路径id(+) And a.Id = c.路径id(+) And  C.版本号=(Select max(版本号) from 临床路径版本 D where A.id=d.路径ID(+)) And C.停用时间 Is NULL "
    strSql = strSql & strFilter
    If InStr(mstrPrivs, "全院路径") = 0 Then
        '没有权限时，只能对只应用于本科的路径进行处理
        strSql = strSql & _
                 " And A.通用 = 2 And Exists" & vbNewLine & _
                 "      (Select 1 From 部门人员 C,临床路径科室 D  " & vbNewLine & _
                 "       Where C.人员id = [1] And D.科室id = C.部门id And 路径id = A.ID)"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)

    '记录现在选中的反馈
    If rptPath.SelectedRows.count > 0 Then
        If Not rptPath.SelectedRows(0).GroupRow Then
            lngPreIdx = rptPath.SelectedRows(0).Index    '用于快速重新定位
            lngPreID = rptPath.SelectedRows(0).Record(COL_ID).Value
        End If
    End If
    
    rptPath.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptPath.Records.Add()
        Set objItem = objRecord.AddItem(Val(rsTmp!ID))
        Set objItem = objRecord.AddItem("")
        If Val(rsTmp!性质 & "") = 1 Then
            objItem.Icon = ilsPic.ListImages("Merge").Index - 1
        Else
            objItem.Icon = ilsPic.ListImages("Path").Index - 1
        End If
        Set objItem = objRecord.AddItem("")
        If NVL(rsTmp!是否存在分支, 0) <> 0 Then
            objItem.Icon = ilsPic.ListImages("branch").Index - 1
        End If
        Set objItem = objRecord.AddItem("")
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!分类, "<未指定分类>")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!编码)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!名称)))
        intType = Val(rsTmp!审核状态 & "")
        If tbcPath.Selected.Tag = "审核通过" Then
            If Val(rsTmp!已审核) = 1 And intType = 0 Then '历史数据处理
                intType = 4
            End If
        End If
        Set objItem = objRecord.AddItem(intType)
        Select Case tbcPath.Selected.Tag
        Case "待审核"
            If gbln双审核 Then
                strTmp = IIf(intType = 1, "待药剂科审核", "待医务科审核")
            Else
                strTmp = "待审核"
            End If
        Case "审核通过"
            If gbln双审核 Then
                strTmp = IIf(intType = 2, "药剂科审核通过", "医务科审核通过")
            Else
                strTmp = "审核通过"
            End If
        Case "审核未过"
            If gbln双审核 Then
                strTmp = IIf(intType = 3, "药剂科审核未过", "医务科审核未过")
            Else
                strTmp = "审核未过"
            End If
        End Select
        objItem.Caption = strTmp

        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!病例分型)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!适用病情)))
        Set objItem = objRecord.AddItem(CStr(Decode(NVL(rsTmp!适用性别, 0), 0, "", 1, "男", 2, "女")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!适用年龄)))
        Set objItem = objRecord.AddItem(CStr("" & rsTmp!说明))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!通用, 1)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!最新版本, 0)))
        
        If gbln双审核 And InStr(mstrPrivs, ";药剂科审核;") > 0 And InStr(mstrPrivs, ";审核;") > 0 Then
            If tbcPath.Selected.Tag = "待审核" Then
                Set objItem = objRecord.AddItem(IIf(intType = 1, "药剂科", "医务科"))
            ElseIf tbcPath.Selected.Tag = "审核通过" Then '2,4
                Set objItem = objRecord.AddItem(IIf(intType = 2, "药剂科", "医务科"))
            ElseIf tbcPath.Selected.Tag = "审核未过" Then '3,5
                Set objItem = objRecord.AddItem(IIf(intType = 3, "药剂科", "医务科"))
            End If
        Else
            Set objItem = objRecord.AddItem("")
        End If
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!性质, 0)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!版本号)))
        lngPathColor = IIf(Val(rsTmp!已审核) = 1, vbBlack, &HFF&)
        For i = COL_行号 To COL_通用
            If lngStopColor <> 0 Then
                objRecord.Item(i).ForeColor = lngStopColor
            ElseIf lngPathColor <> 0 Then
                objRecord.Item(i).ForeColor = lngPathColor
            End If
        Next
        rsTmp.MoveNext
    Loop

    rptPath.Populate

    '分类有多个时，显示行号列
    If rptPath.Rows.count - rptPath.Records.count > 1 Then
        rptPath.Columns(COL_行号).Visible = True
        rptPath.Columns(COL_行号).SortAscending = True
    Else
        rptPath.Columns(COL_行号).Visible = False
    End If

    '行号赋值
    For i = 0 To rptPath.Rows.count - 1
        With rptPath.Rows(i)
            If .GroupRow Then intTypeNum = intTypeNum + 1
            If Not .GroupRow Then
                .Record(COL_行号).Value = i - intTypeNum + 1
            End If
        End With
    Next

    If rptPath.Rows.count = 0 Then
        Call ClearSubData
    Else
        If str分类 <> "" And str编码 <> "" Then
            For i = 0 To rptPath.Rows.count - 1
                If Not rptPath.Rows(i).GroupRow Then
                    If rptPath.Rows(i).Record(COL_分类).Value = str分类 _
                       And rptPath.Rows(i).Record(COL_编码).Value = str编码 Then
                        Set objRow = rptPath.Rows(i): Exit For
                    End If
                End If
            Next
        Else
            If lngPreID <> 0 Then
                '先快速定位
                If lngPreIdx <= rptPath.Rows.count - 1 Then
                    If Not rptPath.Rows(lngPreIdx).GroupRow Then
                        If rptPath.Rows(lngPreIdx).Record(COL_ID).Value = lngPreID Then
                            Set objRow = rptPath.Rows(lngPreIdx)
                        End If
                    End If
                End If
                '再进行查找
                If objRow Is Nothing Then
                    For i = 0 To rptPath.Rows.count - 1
                        If Not rptPath.Rows(i).GroupRow Then
                            If rptPath.Rows(i).Record(COL_ID).Value = lngPreID Then
                                Set objRow = rptPath.Rows(i): Exit For
                            End If
                        End If
                    Next
                End If
            End If
            '取第一个非分组行
            If objRow Is Nothing Then
                For i = 0 To rptPath.Rows.count - 1
                    If Not rptPath.Rows(i).GroupRow Then Set objRow = rptPath.Rows(i): Exit For
                Next
            End If
        End If

        Set rptPath.FocusedRow = objRow    '该行选中且显示在可见区域,并引发SelectionChanged事件
        Me.staThis.Panels(2).Text = "共有 " & rptPath.Records.count & " 个临床路径"
    End If

    Screen.MousePointer = 0
    RefreshData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadAduit(ByVal lngPathID As Long, ByVal lngVersion As Long)

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim objRecord   As ReportRecord
    Dim objItem     As ReportRecordItem
    If lngPathID = 0 Then
        rptLog.Records.DeleteAll
        rptLog.Populate
        Exit Sub
    End If
    If gbln双审核 Then
        strSql = "Select Decode(操作状态, 1, '医务科审核通过', 2, '医务科审核未过', 3, '药剂科审核通过', 4, '药剂科审核未过') As 操作状态, NVL(操作说明,'未填写') As 操作说明, 操作人员, 操作时间" & vbNewLine & _
            "From 临床路径审核" & vbNewLine & _
            "Where 路径id = [1] And 版本号 = [2]" & vbNewLine & _
            "Order By 操作时间 Desc"
    Else
        strSql = "Select Decode(操作状态, 1, '审核通过', 2, '审核未过', 3, '药剂科审核通过', 4, '药剂科审核未过') As 操作状态, NVL(操作说明,'未填写') As 操作说明, 操作人员, 操作时间" & vbNewLine & _
            "From 临床路径审核" & vbNewLine & _
            "Where 路径id = [1] And 版本号 = [2]" & vbNewLine & _
            "Order By 操作时间 Desc"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, lngVersion)
    
    rptLog.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptLog.Records.Add()
        Set objItem = objRecord.AddItem(rsTmp!操作状态 & "")
        Set objItem = objRecord.AddItem(rsTmp!操作说明 & "")
        Set objItem = objRecord.AddItem(rsTmp!操作人员 & "")
        Set objItem = objRecord.AddItem(Format(rsTmp!操作时间 & "", "YYYY-MM-DD HH:MM:SS"))
        rsTmp.MoveNext
    Loop

    rptLog.Populate
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearSubData()
    Dim i As Integer

    lbl说明.Caption = "说明："

    lbl科室(1).Caption = ""

    vsgIllness.Rows = 0
    vsgIllness.Rows = 5

    Me.staThis.Panels(2).Text = ""
    mlng路径ID = 0
    mlng版本号 = 0
    Call mfrmContent.zlRefresh(0, mstrPrivs, lbl科室(1).Caption, vsgIllness.Tag, 2, 0)
    Call LoadAduit(0, 0)
    Call Form_Resize
    Call picList_Resize
End Sub

Private Sub tbcPath_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Me.Visible Then
        Call RefreshData
    End If
End Sub

Private Sub InitVsgIllness()
'功能:初始化对应病种

    With vsgIllness
        .Cols = 3
        .Rows = 5
        .FixedCols = 0
        .FixedRows = 0
        .AllowSelection = False
        .BackColorBkg = vbWhite
        .RowHeightMin = 300
        .Appearance = flexXPThemes
        .BorderStyle = flexBorderNone
        .ScrollBars = flexScrollBarVertical
        .GridLines = flexGridNone
        .ColWidthMin = .Width / .Cols
    End With
End Sub

Private Sub vsgIllness_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vsgIllness.FontUnderline = True
    vsgIllness.ForeColor = RGB(0, 0, 128)
    vsgIllness.ToolTipText = vsgIllness.Text
End Sub

