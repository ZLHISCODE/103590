VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMedRatio 
   Caption         =   "药占比查询"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11145
   Icon            =   "frmMedRatio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   11145
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picQuery 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   2520
      ScaleHeight     =   2400
      ScaleWidth      =   7620
      TabIndex        =   3
      Top             =   690
      Width           =   7620
      Begin VSFlex8Ctl.VSFlexGrid vsQuery 
         Height          =   2295
         Left            =   45
         TabIndex        =   4
         Top             =   60
         Width           =   7530
         _cx             =   13282
         _cy             =   4048
         Appearance      =   1
         BorderStyle     =   1
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
         BackColorSel    =   16764057
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   8000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMedRatio.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
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
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   3750
      ScaleHeight     =   675
      ScaleWidth      =   1440
      TabIndex        =   1
      Top             =   4380
      Width           =   1470
   End
   Begin VB.Frame fraAdviceUD 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   2550
      MousePointer    =   7  'Size N S
      TabIndex        =   0
      Top             =   3210
      Width           =   7815
   End
   Begin XtremeSuiteControls.TabControl tbcDraw 
      Height          =   1515
      Left            =   2535
      TabIndex        =   2
      Top             =   3780
      Width           =   7590
      _Version        =   589884
      _ExtentX        =   13388
      _ExtentY        =   2672
      _StockProps     =   64
   End
   Begin XtremeSuiteControls.TabControl tbcQuery 
      Height          =   3435
      Left            =   2445
      TabIndex        =   5
      Top             =   165
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   6059
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   6330
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16748
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Text            =   "通过"
            TextSave        =   "通过"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Text            =   "疑问"
            TextSave        =   "疑问"
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
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   1755
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   270
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMedRatio.frx":6868
      Left            =   915
      Top             =   285
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMedRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'查询条件
Private Type QUERY_COND
    Drug As Boolean
    FeeRan As Byte '费用范围 0-全院，1-门诊，2-住院
    CouWay As Byte  '统计方式 0-科室，1-开单人，2-病人
    FilIDs As String
    DateBegin As Date
    dateEnd As Date
End Type
Private mvQuery As QUERY_COND
Private WithEvents mfrmCond As frmMedRatioCond
Attribute mfrmCond.VB_VarHelpID = -1
Private mvarColor As Variant
Private Const PI  As Double = 3.1415926
Private mstrWay As String
Private mstrPrivs As String
Private mblnCanSave As Boolean

Public Sub ShowMe(frmParent As Object, strPriv As String)
    mstrPrivs = strPriv
    Me.Show , frmParent
End Sub

Private Sub Form_Load()
    '图形颜色
    mvarColor = Array()
    ReDim mvarColor(3)
    mvarColor(0) = &H8080FF
    mvarColor(1) = &H80FF80
    mvarColor(2) = &HFFFF00
    mvarColor(3) = &HC0C0&
    
    Call MainDefCommandBar
    
    Call InitFilterForm
    
    Me.WindowState = vbMaximized
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到Excel")
            objControl.IconId = conMenu_Edit_NextPage
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_SaveJpeg, "图片另存为")
            objControl.IconId = conMenu_Edit_Save
        Set objControl = .Add(xtpControlButton, conMenu_File_MedRecSetup, "设置")
            objControl.IconId = conMenu_File_Parameter
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh
    End With
    
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitFilterForm()
'功能：主窗口左边部分的条件窗体
    Dim objItem As TabControlItem
    Dim objPane As Pane
    
    '条件栏----------------------------------------------
    Set mfrmCond = New frmMedRatioCond
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 250, 400, DockLeftOf, Nothing)
    objPane.Title = "查询条件"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    With Me.tbcQuery
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
        End With
        Set objItem = .InsertItem(0, "分类统计", picQuery.hwnd, 0): objItem.Color = tbcQuery.PaintManager.ColorSet.ButtonNormal
        Set objItem = .InsertItem(1, "抗菌药物", picQuery.hwnd, 0): objItem.Color = &HC0C0FF
        Set objItem = .InsertItem(2, "基本药物", picQuery.hwnd, 0): objItem.Color = &HAC0FF
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    
    With Me.tbcDraw
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
        End With
        Set objItem = .InsertItem(0, "饼图", PicDraw.hwnd, 0): objItem.Color = &HC0C0FF
        Set objItem = .InsertItem(1, "直方图", PicDraw.hwnd, 0): objItem.Color = &HC0C0FF
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    tbcDraw.Tag = "Pie"
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_MedRecSetup
            If InStr(";" & mstrPrivs & ";", ";参数设置;") = 0 Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
        Case conMenu_File_SaveJpeg, conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel
            Control.Enabled = mblnCanSave
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    With Me.tbcQuery
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = (lngBottom - lngTop) / 3
    End With
        
    fraAdviceUD.Left = tbcQuery.Left
    fraAdviceUD.Top = tbcQuery.Top + tbcQuery.Height - 45
    fraAdviceUD.Width = tbcQuery.Width
    
    With tbcDraw
        .Left = lngLeft
        .Top = fraAdviceUD.Top
        .Width = lngRight - lngLeft
        .Height = (lngBottom - lngTop) * 2 / 3
    End With
End Sub

Private Sub fraAdviceUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tbcQuery.Height + Y < 1000 Or tbcDraw.Height - Y < 6000 Then Exit Sub
        fraAdviceUD.Top = fraAdviceUD.Top + Y
        tbcQuery.Height = tbcQuery.Height + Y
        tbcDraw.Top = tbcDraw.Top + Y
        tbcDraw.Height = tbcDraw.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_MedRecSetup
            frmMedRatioFilter.Show 1, Me
        Case conMenu_View_Refresh
            Select Case tbcDraw.Tag
                Case "Pie"
                    Call DrawPie(vsQuery.RowSel)
                Case "Bar"
                    Call DrawBar(vsQuery.RowSel)
            End Select
        Case conMenu_File_SaveJpeg
            Call SavePic
        Case conMenu_File_Print
            Call OutputData(1)
        Case conMenu_File_Preview
            Call OutputData(2)
        Case conMenu_File_Excel
            Call OutputData(3)
        Case conMenu_Help_Help
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Item.Handle = mfrmCond.hwnd
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmCond Is Nothing Then Unload mfrmCond
    Set mfrmCond = Nothing
    mblnCanSave = False
End Sub

Private Sub mfrmCond_CountWay(ByVal strWay As String, ByVal blnDrug As Boolean)
    mstrWay = strWay
    mvQuery.Drug = blnDrug
    Call InitAdviceTable(tbcQuery.Selected.Index, strWay, blnDrug)
End Sub

Private Function LoadQueryData() As Boolean
'功能：查询数据
    Dim rsTmp As ADODB.Recordset
    Dim rsFilter As ADODB.Recordset
    Dim strSQL As String
    Dim strSQLFilter As String
    Dim strFeePro As String '费用性质 结帐金额 实收金额
    Dim strTmp, strWay As String
    Dim strGroup As String
    Dim dblRatio1, dblRatio2, dblRatio3, dblRatio4 As Double
    Dim dblTmp As Double
    Dim i As Integer
    Dim strSQL1, strSQL2, strType As String
    
    Screen.MousePointer = 11
    If mvQuery.CouWay = 0 Then '科室
        strWay = " c.名称 As 名称,"
        strTmp = " Table(f_Num2list([1])) B, 部门表 C Where a.开单部门id = b.Column_Value And b.Column_Value = c.Id And a.登记时间 Between [2] And [3] And a.记录状态 <> 0"
        strGroup = " Group By c.名称"
    ElseIf mvQuery.CouWay = 1 Then '开单人
        strWay = " c.姓名 As 名称, "
        strTmp = " Table(f_Num2list([1])) B, 人员表 C Where b.Column_Value = c.Id And a.开单人 = c.姓名 And a.登记时间 Between [2] And [3] And a.记录状态 <> 0"
        strGroup = " Group By c.姓名"
    ElseIf mvQuery.CouWay = 2 Then '病人
        strWay = " c.姓名 As 名称,"
        strTmp = " Table(f_Num2list2([1])) B, 病人信息 C Where a.病人id = b.C1 And a.主页id = b.C2 And b.C1 = c.病人id And a.记录状态 <> 0"
        strGroup = " Group By c.姓名"
    End If
    
    strSQLFilter = "Select 内容 From 药比附加条件 Where 类别 = [1] And 内容 Is Not Null"
    Select Case tbcQuery.Selected.Index
        Case 0
            strType = "分类统计"
        Case 1
            strType = "抗菌药物"
        Case 2
            strType = "基本药物"
    End Select
    Set rsFilter = zlDatabase.OpenSQLRecord(strSQLFilter, Me.Caption, strType)
    
    If rsFilter.RecordCount > 0 Then strTmp = strTmp & " And " & rsFilter!内容
    
    Select Case tbcQuery.Selected.Index
        Case 0 '分类统计
            If mvQuery.FeeRan = 0 Then '全院
                strSQL1 = "Select " & strWay & "Decode(a.收费类别, '5', a.实收金额, 0) As 西药费, Decode(a.收费类别, '6', a.实收金额, 0) As 成药费," & vbNewLine & _
                    "Decode(a.收费类别, '7', a.实收金额, 0) As 草药费, Decode(a.收费类别, '5', 0, '6', 0, '7', 0, a.实收金额) As 非药费," & vbNewLine & _
                    "a.实收金额 As 所有费 From 门诊费用记录 A," & strTmp & " And a.记录性质 not in (4,5)"
                strSQL2 = "Select " & strWay & "Decode(a.收费类别, '5', a.实收金额, 0) As 西药费, Decode(a.收费类别, '6', a.实收金额, 0) As 成药费," & vbNewLine & _
                    "Decode(a.收费类别, '7', a.实收金额, 0) As 草药费, Decode(a.收费类别, '5', 0, '6', 0, '7', 0, a.实收金额) As 非药费," & vbNewLine & _
                    "a.实收金额 As 所有费 From 住院费用记录 A," & strTmp
                strSQL1 = IIf(mvQuery.CouWay = 2, Replace(strSQL1, "And a.主页id = b.C2", ""), strSQL1)
                strSQL = strSQL1 & " union all " & strSQL2
                strSQL = "Select /*+ RULE */ 名称, sum(西药费) As 西药费, sum(成药费) As 成药费, sum(草药费) As 草药费, sum(非药费) As 非药费, sum(所有费) As 所有费 From (" & strSQL & ") Having Sum(所有费) > 0  Group By 名称"
            ElseIf mvQuery.FeeRan = 1 Then  '门诊
                strSQL = "Select /*+ RULE */" & strWay & "Sum(Decode(a.收费类别, '5', a.实收金额, 0)) As 西药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '6', a.实收金额, 0)) As 成药费, Sum(Decode(a.收费类别, '7', a.实收金额, 0)) As 草药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '5', 0, '6', 0, '7', 0, a.实收金额)) As 非药费," & vbNewLine & _
                    "Sum(a.实收金额) As 所有费 From 门诊费用记录 A," & strTmp & " And a.记录性质 not in (4,5) Having Sum(a.实收金额) > 0" & strGroup
            ElseIf mvQuery.FeeRan = 2 Then  '住院
                strSQL = "Select /*+ RULE */" & strWay & "Sum(Decode(a.收费类别, '5', a.实收金额, 0)) As 西药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '6', a.实收金额, 0)) As 成药费, Sum(Decode(a.收费类别, '7', a.实收金额, 0)) As 草药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '5', 0, '6', 0, '7', 0, a.实收金额)) As 非药费," & vbNewLine & _
                    "Sum(a.实收金额) As 所有费 From 住院费用记录 A," & strTmp & "  Having Sum(a.实收金额) > 0" & strGroup
            End If
        Case 1 '抗菌药
            If mvQuery.FeeRan = 0 Then '全院
                strSQL1 = "Select " & strWay & "Decode(e.抗生素, 0, 0, a.实收金额) As 抗菌药费, Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0) - Decode(e.抗生素, 0, 0, a.实收金额) As 非抗菌药费," & vbNewLine & _
                    "Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0) As 总药费" & vbNewLine & _
                    "From 门诊费用记录 A,药品规格 D, 药品特性 E, " & strTmp & " And a.收费细目id = d.药品id And d.药名id = e.药名id "
                strSQL2 = "Select " & strWay & "Decode(e.抗生素, 0, 0, a.实收金额) As 抗菌药费, Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0) - Decode(e.抗生素, 0, 0, a.实收金额) As 非抗菌药费," & vbNewLine & _
                    "Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0) As 总药费" & vbNewLine & _
                    "From 住院费用记录 A,药品规格 D, 药品特性 E, " & strTmp & " And a.收费细目id = d.药品id And d.药名id = e.药名id "
                strSQL1 = IIf(mvQuery.CouWay = 2, Replace(strSQL1, "And a.主页id = b.C2", ""), strSQL1)
                strSQL = strSQL1 & " union all " & strSQL2
                strSQL = "Select /*+ RULE */ 名称, sum(抗菌药费) As 抗菌药费, sum(非抗菌药费) As 非抗菌药费, sum(总药费) As 总药费 From (" & strSQL & ") Having Sum(总药费) > 0 Group By 名称"
            ElseIf mvQuery.FeeRan = 1 Then '门诊
                strSQL = "Select /*+ RULE */" & strWay & "Sum(Decode(e.抗生素, 0, 0, a.实收金额)) As 抗菌药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) - Sum(Decode(e.抗生素, 0, 0, a.实收金额)) As 非抗菌药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) As 总药费" & vbNewLine & _
                    "From 门诊费用记录 A,药品规格 D, 药品特性 E, " & strTmp & " And a.收费细目id = d.药品id And d.药名id = e.药名id " & vbNewLine & _
                    "Having Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) > 0" & strGroup
            ElseIf mvQuery.FeeRan = 2 Then '住院
                strSQL = "Select /*+ RULE */" & strWay & "Sum(Decode(e.抗生素, 0, 0, a.实收金额)) As 抗菌药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) - Sum(Decode(e.抗生素, 0, 0, a.实收金额)) As 非抗菌药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) As 总药费" & vbNewLine & _
                    "From 住院费用记录 A,药品规格 D, 药品特性 E, " & strTmp & " And a.收费细目id = d.药品id And d.药名id = e.药名id " & vbNewLine & _
                    "Having Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) > 0" & strGroup
            End If
        Case 2 '基本药
            If mvQuery.FeeRan = 0 Then '全院
                strSQL1 = "Select " & strWay & "Decode(d.基本药物, Null, 0, a.实收金额) As 基本药费, Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0) - Decode(d.基本药物, Null, 0, a.实收金额) As 非基本药费," & vbNewLine & _
                    "Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0) As 总药费" & vbNewLine & _
                    "From 门诊费用记录 A,药品规格 D," & strTmp & " And a.收费细目id = d.药品id "
                strSQL2 = "Select " & strWay & "Decode(d.基本药物, Null, 0, a.实收金额) As 基本药费, Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0) - Decode(d.基本药物, Null, 0, a.实收金额) As 非基本药费," & vbNewLine & _
                    "Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0) As 总药费" & vbNewLine & _
                    "From 住院费用记录 A,药品规格 D," & strTmp & " And a.收费细目id = d.药品id "
                strSQL1 = IIf(mvQuery.CouWay = 2, Replace(strSQL1, "And a.主页id = b.C2", ""), strSQL1)
                strSQL = strSQL1 & " union all " & strSQL2
                strSQL = "Select /*+ RULE */ 名称, sum(基本药费) As 基本药费, sum(非基本药费) As 非基本药费, sum(总药费) As 总药费 From (" & strSQL & ")  Having Sum(总药费) > 0  Group By 名称"
            ElseIf mvQuery.FeeRan = 1 Then '门诊
                strSQL = "Select /*+ RULE */" & strWay & "Sum(Decode(d.基本药物, Null, 0, a.实收金额)) As 基本药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) - Sum(Decode(d.基本药物, Null, 0, a.实收金额)) As 非基本药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) As 总药费" & vbNewLine & _
                    "From 门诊费用记录 A,药品规格 D," & strTmp & " And a.收费细目id = d.药品id " & vbNewLine & _
                    "Having Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) > 0" & strGroup
            ElseIf mvQuery.FeeRan = 2 Then '住院
                strSQL = "Select /*+ RULE */" & strWay & "Sum(Decode(d.基本药物, Null, 0, a.实收金额)) As 基本药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) - Sum(Decode(d.基本药物, Null, 0, a.实收金额)) As 非基本药费," & vbNewLine & _
                    "Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) As 总药费" & vbNewLine & _
                    "From 住院费用记录 A,药品规格 D," & strTmp & " And a.收费细目id = d.药品id " & vbNewLine & _
                    "Having Sum(Decode(a.收费类别, '5', a.实收金额, '6', a.实收金额, '7', a.实收金额, 0)) > 0" & strGroup
            End If
    End Select
    vsQuery.Rows = 2 '清空上一次的数据
    strSQL = strSQL & " Order By 名称"
    On Error GoTo errH
    Call zlCommFun.ShowFlash("正在读取数据，请稍候...")
    Select Case tbcQuery.Selected.Index
        Case 0
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mvQuery.FilIDs, mvQuery.DateBegin, mvQuery.dateEnd)
            If Not rsTmp.EOF Then
                With vsQuery
                    If mvQuery.Drug Then
                        For i = 0 To rsTmp.RecordCount - 1
                            .AddItem ""
                            .TextMatrix(.Rows - 1, 0) = "" & rsTmp!名称
                            .TextMatrix(.Rows - 1, 1) = Format("" & rsTmp!西药费, gstrDec)
                            .TextMatrix(.Rows - 1, 3) = Format("" & rsTmp!成药费, gstrDec)
                            .TextMatrix(.Rows - 1, 5) = Format("" & rsTmp!草药费, gstrDec)
                            .TextMatrix(.Rows - 1, 7) = Format("" & rsTmp!非药费, gstrDec)
                            .TextMatrix(.Rows - 1, 9) = Format("" & rsTmp!所有费, gstrDec)
                            dblRatio1 = Val(rsTmp!西药费) / Val(rsTmp!所有费)
                            dblRatio2 = Val(rsTmp!成药费) / Val(rsTmp!所有费)
                            dblRatio3 = Val(rsTmp!草药费) / Val(rsTmp!所有费)
                            dblRatio4 = 1 - dblRatio1 - dblRatio2 - dblRatio3
                            .TextMatrix(.Rows - 1, 2) = Format(dblRatio1 * 100, "0.0") & "%"
                            .TextMatrix(.Rows - 1, 4) = Format(dblRatio2 * 100, "0.0") & "%"
                            .TextMatrix(.Rows - 1, 6) = Format(dblRatio3 * 100, "0.0") & "%"
                            .TextMatrix(.Rows - 1, 8) = Format(dblRatio4 * 100, "0.0") & "%"
                            rsTmp.MoveNext
                        Next i
                        .Subtotal flexSTSum, -1, 1, "#######" & gstrDec, , vbBlack, False, "总计"
                        .Subtotal flexSTSum, -1, 3, "#######" & gstrDec, , vbBlack, False, "总计"
                        .Subtotal flexSTSum, -1, 5, "#######" & gstrDec, , vbBlack, False, "总计"
                        .Subtotal flexSTSum, -1, 7, "#######" & gstrDec, , vbBlack, False, "总计"
                        .Subtotal flexSTSum, -1, 9, "#######" & gstrDec, , vbBlack, False, "总计"
                        .TextMatrix(2, 0) = "合计"
                        dblRatio1 = Val(.TextMatrix(2, 1)) / Val(.TextMatrix(2, 9))
                        dblRatio2 = Val(.TextMatrix(2, 3)) / Val(.TextMatrix(2, 9))
                        dblRatio3 = Val(.TextMatrix(2, 5)) / Val(.TextMatrix(2, 9))
                        dblRatio4 = 1 - dblRatio1 - dblRatio2 - dblRatio3
                        .TextMatrix(2, 2) = Format(dblRatio1 * 100, "0.0") & "%"
                        .TextMatrix(2, 4) = Format(dblRatio2 * 100, "0.0") & "%"
                        .TextMatrix(2, 6) = Format(dblRatio3 * 100, "0.0") & "%"
                        .TextMatrix(2, 8) = Format(dblRatio4 * 100, "0.0") & "%"
                    Else
                        For i = 0 To rsTmp.RecordCount - 1
                            .AddItem ""
                            .TextMatrix(.Rows - 1, 0) = "" & rsTmp!名称
                            
                            dblTmp = Val("" & rsTmp!所有费) - Val("" & rsTmp!非药费)
                            .TextMatrix(.Rows - 1, 1) = Format(dblTmp, gstrDec)
                            .TextMatrix(.Rows - 1, 3) = Format("" & rsTmp!非药费, gstrDec)
                            .TextMatrix(.Rows - 1, 5) = Format("" & rsTmp!所有费, gstrDec)
                            dblRatio1 = dblTmp / Val(rsTmp!所有费)
                            dblRatio2 = 1 - dblRatio1
                            .TextMatrix(.Rows - 1, 2) = Format(dblRatio1 * 100, "0.0") & "%"
                            .TextMatrix(.Rows - 1, 4) = Format(dblRatio2 * 100, "0.0") & "%"
                            rsTmp.MoveNext
                        Next i
                        .Subtotal flexSTSum, -1, 1, "#######" & gstrDec, , vbBlack, False, "总计"
                        .Subtotal flexSTSum, -1, 3, "#######" & gstrDec, , vbBlack, False, "总计"
                        .Subtotal flexSTSum, -1, 5, "#######" & gstrDec, , vbBlack, False, "总计"
                        .TextMatrix(2, 0) = "合计"
                        dblRatio1 = Val(.TextMatrix(2, 1)) / Val(.TextMatrix(2, 5))
                        dblRatio2 = 1 - dblRatio1
                        .TextMatrix(2, 2) = Format(dblRatio1 * 100, "0.0") & "%"
                        .TextMatrix(2, 4) = Format(dblRatio2 * 100, "0.0") & "%"
                    End If
                End With
            End If
        Case 1
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mvQuery.FilIDs, mvQuery.DateBegin, mvQuery.dateEnd)
            If Not rsTmp.EOF Then
                With vsQuery
                    For i = 0 To rsTmp.RecordCount - 1
                        .AddItem ""
                        .TextMatrix(.Rows - 1, 0) = "" & rsTmp!名称
                        .TextMatrix(.Rows - 1, 1) = Format("" & rsTmp!抗菌药费, gstrDec)
                        .TextMatrix(.Rows - 1, 3) = Format("" & rsTmp!非抗菌药费, gstrDec)
                        .TextMatrix(.Rows - 1, 5) = Format("" & rsTmp!总药费, gstrDec)
                        dblRatio1 = Val(rsTmp!抗菌药费) / Val(rsTmp!总药费)
                        dblRatio2 = 1 - dblRatio1
                        .TextMatrix(.Rows - 1, 2) = Format(dblRatio1 * 100, "0.0") & "%"
                        .TextMatrix(.Rows - 1, 4) = Format(dblRatio2 * 100, "0.0") & "%"
                        rsTmp.MoveNext
                    Next i
                    .Subtotal flexSTSum, -1, 1, "#######" & gstrDec, , vbBlack, False, "总计"
                    .Subtotal flexSTSum, -1, 3, "#######" & gstrDec, , vbBlack, False, "总计"
                    .Subtotal flexSTSum, -1, 5, "#######" & gstrDec, , vbBlack, False, "总计"
                    .TextMatrix(2, 0) = "合计"
                    dblRatio1 = Val(.TextMatrix(2, 1)) / Val(.TextMatrix(2, 5))
                    dblRatio2 = 1 - dblRatio1
                    .TextMatrix(2, 2) = Format(dblRatio1 * 100, "0.0") & "%"
                    .TextMatrix(2, 4) = Format(dblRatio2 * 100, "0.0") & "%"
                End With
            End If
        Case 2
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mvQuery.FilIDs, mvQuery.DateBegin, mvQuery.dateEnd)
            If Not rsTmp.EOF Then
                With vsQuery
                    For i = 0 To rsTmp.RecordCount - 1
                        .AddItem ""
                        .TextMatrix(.Rows - 1, 0) = "" & rsTmp!名称
                        .TextMatrix(.Rows - 1, 1) = Format("" & rsTmp!基本药费, gstrDec)
                        .TextMatrix(.Rows - 1, 3) = Format("" & rsTmp!非基本药费, gstrDec)
                        .TextMatrix(.Rows - 1, 5) = Format("" & rsTmp!总药费, gstrDec)
                        dblRatio1 = Val(rsTmp!基本药费) / Val(rsTmp!总药费)
                        dblRatio2 = 1 - dblRatio1
                        .TextMatrix(.Rows - 1, 2) = Format(dblRatio1 * 100, "0.0") & "%"
                        .TextMatrix(.Rows - 1, 4) = Format(dblRatio2 * 100, "0.0") & "%"
                        rsTmp.MoveNext
                    Next i
                    .Subtotal flexSTSum, -1, 1, "#######" & gstrDec, , vbBlack, False, "总计"
                    .Subtotal flexSTSum, -1, 3, "#######" & gstrDec, , vbBlack, False, "总计"
                    .Subtotal flexSTSum, -1, 5, "#######" & gstrDec, , vbBlack, False, "总计"
                    .TextMatrix(2, 0) = "合计"
                    dblRatio1 = Val(.TextMatrix(2, 1)) / Val(.TextMatrix(2, 5))
                    dblRatio2 = 1 - dblRatio1
                    .TextMatrix(2, 2) = Format(dblRatio1 * 100, "0.0") & "%"
                    .TextMatrix(2, 4) = Format(dblRatio2 * 100, "0.0") & "%"
                End With
            End If
    End Select
    vsQuery.Cell(flexcpAlignment, 0, 0, vsQuery.Rows - 1, 0) = 4
    
    If vsQuery.Rows > 2 Then
        vsQuery.RowSel = 2
        cbsMain.FindControl(, conMenu_View_Refresh, True, True).Execute
    Else
        PicDraw.Cls
        mblnCanSave = False
    End If
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mfrmCond_DoQuery(ByVal bytRan As Byte, ByVal bytWay As Byte, ByVal lngIDs As String, ByVal datBegin As Date, ByVal datEnd As Date)
    mvQuery.FeeRan = bytRan
    mvQuery.CouWay = bytWay
    mvQuery.DateBegin = datBegin
    mvQuery.dateEnd = datEnd
    mvQuery.FilIDs = lngIDs
    Call tbcQuery_SelectedChanged(tbcQuery.Selected)
    vsQuery.SetFocus
End Sub

Private Sub picQuery_Resize()
    With picQuery
        vsQuery.Top = picQuery.ScaleTop
        vsQuery.Left = picQuery.ScaleLeft
        vsQuery.Height = picQuery.ScaleHeight - 40
        vsQuery.Width = picQuery.ScaleWidth
    End With
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub DrawPie(ByVal Row As Integer)
'功能：画饼图  构成分析
    Dim varRatio As Variant
    Dim varInfo As Variant
    Dim strName As String
    Dim i  As Integer
    
    PicDraw.Cls
    mblnCanSave = False
    If Row < 2 Then Exit Sub
    
    varRatio = Array()
    varInfo = Array()
    
    '获取图例显示文字和比例数据
    With vsQuery
        strName = .TextMatrix(Row, 0) & "   费用比例"
        For i = 1 To .Cols - 1
            If InStr(.TextMatrix(Row, i), "%") > 0 Then
                ReDim Preserve varRatio(UBound(varRatio) + 1)
                ReDim Preserve varInfo(UBound(varInfo) + 1)
                
                varRatio(UBound(varRatio)) = Val(Mid(.TextMatrix(Row, i), 1, Len(.TextMatrix(Row, i)) - 1)) / 100
                If UBound(varRatio) > 0 Then
                    varRatio(UBound(varRatio)) = varRatio(UBound(varRatio)) + varRatio(UBound(varRatio) - 1)
                End If
                
                varInfo(UBound(varInfo)) = .TextMatrix(0, i) & " " & .TextMatrix(Row, i)
            End If
        Next i
        varRatio(UBound(varRatio)) = 1
    End With
    
    PicDraw.AutoRedraw = True  '设置自动绘图
    PicDraw.Cls                '清空绘图区域
    PicDraw.DrawWidth = 1
    PicDraw.Scale (-1000, 1000)-(1000, -1000) '自定义坐标
    PicDraw.FillStyle = 0    '填充样式为实心
    PicDraw.FillStyle = vbFSTransparent
    PicDraw.FillColor = vbWhite
    '输出标题
    PicDraw.CurrentX = -900: PicDraw.CurrentY = 900
    PicDraw.FontSize = 12
    PicDraw.FontBold = True
    PicDraw.ForeColor = vbRed
    PicDraw.Print strName
    '字体还原
    PicDraw.FontSize = 9
    PicDraw.FontBold = False
    PicDraw.FillStyle = vbFSSolid
    
    '画背景圆
    PicDraw.FillColor = vbBlack
    PicDraw.Circle (0, 0), 500, vbBlack, , , 0.6
    For i = 0 To UBound(varRatio)
        PicDraw.FillColor = mvarColor(i)
        '要考虑特殊情况 比例存在100%的情况，100%的位置，首-中-末
        If varRatio(i) = 0 Then
            i = 1 + i
            PicDraw.FillColor = mvarColor(i)
        End If
        
        If varRatio(i) = 1 And i < UBound(varRatio) Then
            PicDraw.Circle (0, 50), 500, vbBlack, , , 0.6
            Exit For
        End If
                
        If i = 0 Then
            PicDraw.Circle (0, 50), 500, vbBlack, -2 * PI, -2 * PI * varRatio(i), 0.6
        ElseIf varRatio(i) <> varRatio(i - 1) Then
            PicDraw.Circle (0, 50), 500, vbBlack, -2 * PI * IIf(varRatio(i - 1) = 0, 1, varRatio(i - 1)), -2 * PI * varRatio(i), 0.6
        End If
        
        If varRatio(UBound(varRatio) - 1) = 0 Then
            PicDraw.Circle (0, 50), 500, vbBlack, , , 0.6
        End If
    Next i
    
    PicDraw.ForeColor = vbBlack '字体颜色，绘制图例说明
    For i = 0 To UBound(varInfo)
        PicDraw.Line (-900, -300 - i * 150)-(-800, -200 - i * 150), mvarColor(i), BF
        PicDraw.CurrentY = PicDraw.CurrentY - 20
        PicDraw.CurrentX = PicDraw.CurrentX + 20
        PicDraw.Print varInfo(i)
    Next i
    mblnCanSave = True
End Sub

Private Sub DrawBar(ByVal Row As Integer)
'功能：画直方图  构成分析
    Dim varRatio As Variant
    Dim varInfo As Variant
    Dim strName As String
    Dim i As Integer
    
    PicDraw.Cls
    mblnCanSave = False
    
    If Row < 2 Then Exit Sub
    
    varRatio = Array()
    varInfo = Array()
    
    '获取图例显示文字和比例数据
    With vsQuery
        strName = .TextMatrix(Row, 0) & "   费用比例"
        For i = 1 To .Cols - 1
            If InStr(.TextMatrix(Row, i), "%") > 0 Then
                ReDim Preserve varRatio(UBound(varRatio) + 1)
                ReDim Preserve varInfo(UBound(varInfo) + 1)
                varRatio(UBound(varRatio)) = Val(Mid(.TextMatrix(Row, i), 1, Len(.TextMatrix(Row, i)) - 1)) / 100
                varInfo(UBound(varInfo)) = .TextMatrix(0, i) & " " & .TextMatrix(Row, i)
            End If
        Next i
    End With
    
    '开始进行绘图
    PicDraw.AutoRedraw = True  '设置自动绘图
    PicDraw.Cls                '清空绘图区域
    PicDraw.DrawWidth = 1
    PicDraw.Scale (-0.5, 21)-(21, 0) '自定义坐标
    PicDraw.FillStyle = vbFSTransparent
    PicDraw.FillColor = vbWhite
    '输出标题
    PicDraw.CurrentX = 5: PicDraw.CurrentY = 20
    PicDraw.FontSize = 12
    PicDraw.FontBold = True
    PicDraw.ForeColor = vbRed
    PicDraw.Print strName

    '字体还原
    PicDraw.FontSize = 9
    PicDraw.FontBold = False
    PicDraw.ForeColor = vbBlack
    PicDraw.Line (0.2, 3)-(19, 3), vbBlack  '画X轴
    PicDraw.CurrentX = 19.3: PicDraw.CurrentY = PicDraw.CurrentY + 0.3
    PicDraw.Print "类别"
    PicDraw.Line (18.7, 3.3)-(19, 3), vbBlack  '画箭头
    PicDraw.Line (19, 3)-(18.7, 2.7), vbBlack  '画箭头
    
    PicDraw.Line (1, 19)-(1, 0.2), vbBlack  '画Y轴
    PicDraw.CurrentX = 0.7: PicDraw.CurrentY = 20
    PicDraw.Print "比例"
    PicDraw.Line (0.8, 18.3)-(1, 19), vbBlack  '画箭头
    PicDraw.Line (1, 19)-(1.2, 18.3), vbBlack  '画箭头
    For i = 1 To 10  '画销病区人数比例刻度（Y轴）
        PicDraw.Line (0.9, (i * 1.5) + 3)-(1, (i * 1.5) + 3), vbBlack '画Y轴
        PicDraw.CurrentX = 0.4: PicDraw.CurrentY = PicDraw.CurrentY + 0.3
        PicDraw.Print i * 10 & "%"
    Next i
    PicDraw.FillStyle = 0
  
    For i = 0 To UBound(varRatio)
        PicDraw.Line (i * 4.5 + 3.2, 3)-(i * 4.5 + 3.2, 2.99), vbBlack
        PicDraw.CurrentY = PicDraw.CurrentY - 0.3
        PicDraw.FillColor = mvarColor(i)
        PicDraw.Print Split(varInfo(i), " ")(0)
        PicDraw.Line (2.4 + i * 4.5, 15 * varRatio(i) + 3)-(4.4 + i * 4.5, 3), vbBlack, B
        PicDraw.CurrentX = 3 + i * 4.5: PicDraw.CurrentY = 15 * varRatio(i) + 4
        PicDraw.Print Split(varInfo(i), " ")(1)
    Next i
    
    PicDraw.ForeColor = vbBlack '字体颜色，绘制图例说明
    For i = 0 To UBound(varInfo)
        PicDraw.Line (19, 20 - i * 1.5)-(18, 19 - i * 1.5), mvarColor(i), BF
        PicDraw.CurrentY = PicDraw.CurrentY + 0.8
        PicDraw.CurrentX = PicDraw.CurrentX + 1.1
        PicDraw.Print varInfo(i)
    Next i
    mblnCanSave = True
End Sub

Private Sub tbcDraw_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If vsQuery.Rows > 2 Then
        If Item.Index = 0 Then
            Call DrawPie(vsQuery.RowSel)
            tbcDraw.Tag = "Pie"
        ElseIf Item.Index = 1 Then
            Call DrawBar(vsQuery.RowSel)
            tbcDraw.Tag = "Bar"
        End If
    End If
End Sub

Private Sub InitAdviceTable(ByVal intIn As Integer, ByVal strIn As String, ByVal blnDrug As Boolean)
    Dim i As Integer
    
    strIn = IIf(strIn = "", "开单科室", strIn)
    
    Select Case intIn
        Case 0
            With vsQuery
                .Clear
                If blnDrug Then
                    .Rows = 2: .Cols = 10
                    .FixedRows = 2: .FixedCols = 1
                    .MergeRow(0) = True
                    .Cell(flexcpText, 0, 1, 0, 2) = "西药费"
                    .Cell(flexcpText, 0, 3, 0, 4) = "成药费"
                    .Cell(flexcpText, 0, 5, 0, 6) = "草药费"
                    .Cell(flexcpText, 0, 7, 0, 8) = "非药费"
                    For i = 1 To 4
                        .TextMatrix(1, 2 * i - 1) = "金额"
                        .TextMatrix(1, 2 * i) = "比例"
                    Next i
                    .TextMatrix(1, 0) = strIn: .TextMatrix(1, 9) = "总费"
                    .Cell(flexcpAlignment, 0, 0, 1, 9) = 4
                Else
                    .Rows = 2: .Cols = 6
                    .FixedRows = 2: .FixedCols = 1
                    .MergeRow(0) = True
                    .Cell(flexcpText, 0, 1, 0, 2) = "药品费"
                    .Cell(flexcpText, 0, 3, 0, 4) = "非药费"
                    For i = 1 To 2
                        .TextMatrix(1, 2 * i - 1) = "金额"
                        .TextMatrix(1, 2 * i) = "比例"
                    Next i
                    .TextMatrix(1, 0) = strIn: .TextMatrix(1, 5) = "总费"
                    .Cell(flexcpAlignment, 0, 0, 1, 5) = 4
                End If
            End With
        Case 1
            With vsQuery
                .Clear
                .Rows = 2: .Cols = 6
                .FixedRows = 2: .FixedCols = 1
                .MergeRow(0) = True
                .Cell(flexcpText, 0, 1, 0, 2) = "抗菌药费"
                .Cell(flexcpText, 0, 3, 0, 4) = "非抗菌药费"
                For i = 1 To 2
                    .TextMatrix(1, 2 * i - 1) = "金额"
                    .TextMatrix(1, 2 * i) = "比例"
                Next i
                .TextMatrix(1, 0) = strIn: .TextMatrix(1, 5) = "药费"
                .Cell(flexcpAlignment, 0, 0, 1, 5) = 4
            End With
        Case 2
            With vsQuery
                .Clear
                .Rows = 2: .Cols = 6
                .FixedRows = 2: .FixedCols = 1
                .MergeRow(0) = True
                .Cell(flexcpText, 0, 1, 0, 2) = "基本药费"
                .Cell(flexcpText, 0, 3, 0, 4) = "非基本药费"
                For i = 1 To 2
                    .TextMatrix(1, 2 * i - 1) = "金额"
                    .TextMatrix(1, 2 * i) = "比例"
                Next i
                .TextMatrix(1, 0) = strIn: .TextMatrix(1, 5) = "药费"
                .Cell(flexcpAlignment, 0, 0, 1, 5) = 4
            End With
    End Select
    PicDraw.Cls
End Sub

Private Sub tbcQuery_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'选项卡切换时，更改表格的表头
    PicDraw.Cls
    Call InitAdviceTable(Item.Index, mstrWay, mvQuery.Drug)
    If Visible Then Call LoadQueryData
End Sub

Private Sub vsQuery_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow <= 1 Then Exit Sub
    Select Case tbcDraw.Tag
        Case "Pie"
            Call DrawPie(NewRow)
        Case "Bar"
            Call DrawBar(NewRow)
    End Select
End Sub

Private Sub SavePic()
'功能：另存为图片
    dlgFile.Filter = "Jpeg|*.jpg|Bmp|*.bmp|Icon|*.ico|Png|*.png"
    dlgFile.FileName = "图" & Format(Now, "yyyymmddhhmmss")
    Call dlgFile.ShowSave
    If InStr(dlgFile.FileName, ":") <> 2 Then Exit Sub
    PicDraw.AutoRedraw = True
    SavePicture PicDraw.Image, dlgFile.FileName
End Sub

Private Sub OutputData(ByVal bytIn As Byte)
'功能：打印输出
'参数：bytIn  1-打印,2-预览,3-输出到EXCEL
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    
    objOut.Title.Text = Me.Caption
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    Set objRow = New zlTabAppRow
    objRow.Add "时间范围： " & mvQuery.DateBegin & " 到 " & mvQuery.dateEnd
    objOut.UnderAppRows.Add objRow
    
    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm")
    objOut.BelowAppRows.Add objRow
    Set objOut.Body = vsQuery
    zlPrintOrView1Grd objOut, bytIn
End Sub


