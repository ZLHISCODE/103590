VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStPathMain 
   Caption         =   "标准路径参考"
   ClientHeight    =   9435
   ClientLeft      =   3240
   ClientTop       =   1395
   ClientWidth     =   12765
   Icon            =   "frmStPathMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   12765
   WindowState     =   2  'Maximized
   Begin XtremeReportControl.ReportControl rptStPath 
      Height          =   6495
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   5100
      _Version        =   589884
      _ExtentX        =   8996
      _ExtentY        =   11456
      _StockProps     =   0
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "查找病人(Ctrl+F)"
      Top             =   0
      Width           =   1155
   End
   Begin VB.PictureBox picStPathDetial 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   5300
      ScaleHeight     =   7095
      ScaleWidth      =   7575
      TabIndex        =   1
      Top             =   270
      Width           =   7575
      Begin VB.PictureBox picPathTable 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   720
         ScaleHeight     =   2295
         ScaleWidth      =   6495
         TabIndex        =   3
         Top             =   360
         Width           =   6495
         Begin VB.Frame fraTableTile 
            BackColor       =   &H00F0F4E4&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   6255
            Begin VB.Label lblTableTile 
               AutoSize        =   -1  'True
               BackColor       =   &H00F0F4E4&
               Height          =   180
               Left            =   120
               TabIndex        =   7
               Top             =   0
               Width           =   90
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPathTable 
            Height          =   975
            Left            =   0
            TabIndex        =   4
            Top             =   1320
            Width           =   3585
            _cx             =   6324
            _cy             =   1720
            Appearance      =   0
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorAlternate=   16777215
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   3
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   3
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   20
            RowHeightMax    =   5000
            ColWidthMin     =   100
            ColWidthMax     =   12000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmStPathMain.frx":058A
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
      End
      Begin XtremeSuiteControls.TabControl tbcStPath 
         Height          =   7335
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   6615
         _Version        =   589884
         _ExtentX        =   11668
         _ExtentY        =   12938
         _StockProps     =   64
      End
      Begin VB.PictureBox picPathCourse 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   240
         ScaleHeight     =   3735
         ScaleWidth      =   6015
         TabIndex        =   8
         Top             =   3000
         Width           =   6015
         Begin RichTextLib.RichTextBox rtfPathCourse 
            Height          =   4095
            Left            =   120
            TabIndex        =   9
            Top             =   60
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   7223
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmStPathMain.frx":05F6
         End
      End
   End
   Begin VB.Frame fraSplit 
      Caption         =   "Frame1"
      Height          =   7335
      Left            =   5200
      MousePointer    =   9  'Size W E
      TabIndex        =   0
      Top             =   0
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTemp 
      Height          =   900
      Left            =   7200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
   Begin XtremeSuiteControls.TabControl tbcPathName 
      Height          =   975
      Left            =   480
      TabIndex        =   13
      Top             =   720
      Width           =   2535
      _Version        =   589884
      _ExtentX        =   4471
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblFind 
      Caption         =   "路径名称"
      Height          =   220
      Left            =   1440
      TabIndex        =   11
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmStPathMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrs表单 As New ADODB.Recordset '标准路径表单内容
Private mrs表头信息 As New ADODB.Recordset '标准路径表单的表头以及阶段数、分类数目等信息
Private mlngStPathID As Long '选中的标准路径的ID
Private Const M_INT_STEPNUM = 3 '固定显示阶段数
Private mstrTilePos As String '标准路径流程中路径流程段落的开始位置，格式为：段落1开始位置（0）,段落2，段落3

'列枚举
Private Enum PathListCols
    COL_ID = 0
    COL_科室名称 = 1
    COL_编码 = 2
    COL_路径名称 = 3
    COL_版本说明 = 4
    COL_疾病编码 = 5
    COL_手术编码 = 6
End Enum
'当前获得焦点的控件枚举
Private Enum FocusContrl
    FC_PathList = 0
    FC_Course = 1
    FC_TBC = 2
    FC_Table = 3
End Enum
Private mintFocus As Integer '当前获得焦点的控件

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long

    Select Case Control.ID
        Case conMenu_File_PrintSet
            Call zlPrintSet
        Case conMenu_File_Preview
            Call zlRptPrint(0)
        Case conMenu_File_Print
            Call zlRptPrint(1)
        Case conMenu_File_Excel
            Call zlRptPrint(3)
        Case conMenu_View_ToolBar_Button '工具栏
            For i = 2 To cbsMain.Count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text '按钮文字
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
                        objControl.Style = xtpButtonIcon
                    Else
                        objControl.Style = IIf(Control.Checked, xtpButtonIconAndCaption, xtpButtonIcon)
                    End If
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size '大图标
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case conMenu_View_Expend_CurCollapse '折叠当前组
            If rptStPath.SelectedRows.Count > 0 Then
                If rptStPath.SelectedRows(0).GroupRow Then
                    rptStPath.SelectedRows(0).Expanded = False
                ElseIf Not rptStPath.SelectedRows(0).ParentRow Is Nothing Then
                    If rptStPath.SelectedRows(0).ParentRow.GroupRow Then
                        rptStPath.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '因折叠定位到分组上,不会自动激活该事件
            Call rptStPath_SelectionChanged
        Case conMenu_View_Expend_CurExpend '展开当前组
            If rptStPath.SelectedRows.Count > 0 Then
                rptStPath.SelectedRows(0).Expanded = True
            End If
        Case conMenu_View_Expend_AllCollapse '折叠所有组
            For Each objRow In rptStPath.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '因折叠定位到分组上,不会自动激活该事件
            Call rptStPath_SelectionChanged
        Case conMenu_View_Expend_AllExpend '展开所有组
            For Each objRow In rptStPath.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        Case conMenu_View_Find '查找
            If Me.ActiveControl Is txtFind Then
                txtFind.SetFocus '有时需要定位一下
                If txtFind.Text <> "" Then
                    Call FindPath(False)
                End If
            Else
                txtFind.SetFocus
            End If
        Case conMenu_View_FindNext '查找下一个
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call FindPath(True)
            End If
        '路径目录增删改
        Case conMenu_Edit_NewPath
            Call InsertItem(0)
        Case conMenu_Edit_ModifyPath
            Call ModItem(0)
        Case conMenu_Edit_DelPath
            Call DeleteItem(0, Val(rptStPath.SelectedRows(0).Record(COL_ID).Value))
        '路径流程段落增删改
        Case conMenu_Edit_NewCourseItem
            Call InsertItem(1)
        Case conMenu_Edit_ModifyCourseItem
            Call ModItem(1)
        Case conMenu_Edit_DelCourseItem
            Call DeleteItem(1, Val(rptStPath.SelectedRows(0).Record(COL_ID).Value))
        '路径表单增删改
        Case conMenu_Edit_NewTable
            Call InsertItem(2)
        Case conMenu_Edit_ModifyTable
            Call ModItem(2)
        Case conMenu_Edit_DelTable
            Call DeleteItem(2, Val(rptStPath.SelectedRows(0).Record(COL_ID).Value))
        '表单内容修改
        Case conMenu_Edit_ModifyTableContent
            Call ModItem(3)
            
        Case conMenu_Help_Web_Home 'Web上的中联
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '发送反馈
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 10))
        Case conMenu_File_Exit '退出
            Unload Me
        Case conMenu_File_ImportPathTable
            If frmImportPath.ShowMe(Me, mlngStPathID) = True Then
                Call LoadStPathList(tbcPathName.Selected.Index)
            End If
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngTop As Long, lngLeft As Long, lngBottom As Long, lngRight As Long
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    tbcPathName.Left = lngLeft
    tbcPathName.Top = lngTop
    tbcPathName.Height = lngBottom - lngTop
    tbcPathName.Width = (lngRight - lngLeft) * 0.3
    
    fraSplit.Top = lngTop
    fraSplit.Left = tbcPathName.Left + tbcPathName.Width + 30
    fraSplit.Height = lngBottom - lngTop
    
    picStPathDetial.Top = lngTop
    picStPathDetial.Left = fraSplit.Left + fraSplit.Width + 30
    picStPathDetial.Width = lngRight - picStPathDetial.Left
    picStPathDetial.Height = lngBottom - lngTop
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean, blnParent As Boolean
    
    '根据当前活动的控件获得可用性
    blnEnabled = True
    
    Select Case mintFocus
        Case FC_Course
            If rtfPathCourse.SelStart >= Len(rtfPathCourse.Text) And Len(rtfPathCourse.Text) <> 0 Or Len(rtfPathCourse.Text) = 0 Then
                blnEnabled = False
            End If
        Case FC_TBC
            If tbcStPath.Selected.Index = 0 Then blnEnabled = False
        Case FC_PathList
            blnEnabled = True
        Case FC_Table
            If tbcStPath.Selected.Index = 0 Then blnEnabled = False
    End Select
    If rptStPath.Rows.Count <> 0 Then
        If rptStPath.SelectedRows.Count <> 0 Then
            blnParent = rptStPath.SelectedRows(0).GroupRow
        End If
    Else
        blnParent = True
    End If
    
    blnEnabled = blnEnabled And Not blnParent
    
    Select Case Control.ID
        '路径目录增删改
        Case conMenu_Edit_NewPath
            Control.Enabled = True
        Case conMenu_Edit_ModifyPath
            Control.Enabled = Not blnParent
        Case conMenu_Edit_DelPath
            Control.Enabled = Not blnParent
        '路径流程段落增删改
        Case conMenu_Edit_NewCourseItem
            Control.Enabled = Not blnParent
        Case conMenu_Edit_ModifyCourseItem
            Control.Enabled = (mintFocus = FC_Course) And blnEnabled
        Case conMenu_Edit_DelCourseItem
            Control.Enabled = (mintFocus = FC_Course) And blnEnabled
        '路径表单增删改
        Case conMenu_Edit_NewTable
            Control.Enabled = Not blnParent
        Case conMenu_Edit_ModifyTable
            Control.Enabled = (mintFocus = FC_Table Or mintFocus = FC_TBC) And blnEnabled
        Case conMenu_Edit_DelTable
            Control.Enabled = (mintFocus = FC_Table Or mintFocus = FC_TBC) And blnEnabled And tbcStPath.ItemCount > 2
        '表单内容修改
        Case conMenu_Edit_ModifyTableContent
            Control.Enabled = (mintFocus = FC_Table Or mintFocus = FC_TBC) And blnEnabled
        Case conMenu_View_ToolBar_Button '工具栏
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Size '大图标
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_Expend_CurExpend '展开当前组
            blnEnabled = False
            If rptStPath.SelectedRows.Count > 0 Then
                If rptStPath.SelectedRows(0).GroupRow Then
                    blnEnabled = Not rptStPath.SelectedRows(0).Expanded
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend_CurCollapse '折叠当前组
            blnEnabled = False
            If rptStPath.SelectedRows.Count > 0 Then
                If rptStPath.SelectedRows(0).GroupRow Then
                    blnEnabled = rptStPath.SelectedRows(0).Expanded
                ElseIf Not rptStPath.SelectedRows(0).ParentRow Is Nothing Then
                    If rptStPath.SelectedRows(0).ParentRow.GroupRow Then
                        blnEnabled = rptStPath.SelectedRows(0).ParentRow.Expanded
                    End If
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend '折叠/展开组
            Control.Enabled = rptStPath.GroupsOrder.Count > 0 And rptStPath.Rows.Count > 0
    End Select
End Sub

Private Sub Form_Load()

    mlngStPathID = 0
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = False
        .ShowTextBelowIcons = False
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization True
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    'tbcPathName路径参考
    With Me.tbcPathName
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem 0, "西医参考", rptStPath.hwnd, 0
        .InsertItem 1, "中医参考", rptStPath.hwnd, 0
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    '初始化tbcControl
    With tbcStPath
    
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        
        .AllowReorder = False
        '初次加载数据只加载选项卡以及标准住院流程
        Call .InsertItem(0, "标准住院流程", picPathCourse.hwnd, 0)
        .Item(0).Selected = True '默认选择标准住院流程
        
    End With
    Call MainDefCommandBar
    '初始化标准路径列表
    Call InitPathList
    '加载标准路径目录
    Call LoadStPathList(tbcPathName.Selected.Index)
End Sub

Private Sub Form_Resize()
'功能：设置tbcPathNameh与picStPathDetial的位置大小
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        Me.Height = IIf(Me.Height < 9000, 9000, Me.Height)
        Me.Width = IIf(Me.Width < 12000, 12000, Me.Width)
    End If
    Call cbsMain_Resize
End Sub


Private Sub Form_Unload(Cancel As Integer)
'功能：清除模块级变量值
    Set mrs表单 = Nothing
    Set mrs表头信息 = Nothing
    mlngStPathID = 0
    mstrTilePos = ""
    mintFocus = -1
    
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'功能：实现标准路径清单与标准路径内容自由拖动大小

    If Button = 1 Then
        If tbcPathName.Width + x > 11000 Or tbcPathName.Width + x < 2000 Then Exit Sub
        
        fraSplit.Left = fraSplit.Left + x
        tbcPathName.Width = fraSplit.Left - 30 - tbcPathName.Left
        
        picStPathDetial.Left = fraSplit.Left + 30
        picStPathDetial.Width = Me.ScaleWidth - picStPathDetial.Left
        
        Me.Refresh
    End If
    
End Sub

Private Sub picPathCourse_Resize()
'功能：实现标准路径流程内容的大小设置

    rtfPathCourse.Width = picPathCourse.Width - rtfPathCourse.Left - 120
    rtfPathCourse.Height = picPathCourse.Height - rtfPathCourse.Top
    
End Sub

Private Sub picPathTable_Resize()
'功能：设置表单表头与表单所在控件的位置与大小

    fraTableTile.Height = lblTableTile.Height + 60
    fraTableTile.Width = picPathTable.Width
    lblTableTile.Width = fraTableTile.Width
    lblTableTile.Width = fraTableTile.Width - lblTableTile.Left
    
    vsPathTable.Top = fraTableTile.Top + fraTableTile.Height
    vsPathTable.Height = picPathTable.Height - vsPathTable.Top
    vsPathTable.Width = picPathTable.Width - vsPathTable.Left
    
End Sub

Private Sub picStPathDetial_Resize()
'功能：标准路径内容区的大小设置

    tbcStPath.Top = 0
    tbcStPath.Left = 0
    tbcStPath.Width = picStPathDetial.Width
    tbcStPath.Height = picStPathDetial.Height
    picPathTable.Width = tbcStPath.Width
    picPathCourse.Width = tbcStPath.Width
    
End Sub

Private Sub rptStPath_GotFocus()
    mintFocus = FC_PathList
End Sub

Private Sub rptStPath_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBar
    
    mintFocus = FC_PathList

    If Button = 2 Then
        Set objHitTest = rptStPath.HitTest(x, y)
        If objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
            With objPopup.Controls
                 .Add xtpControlButton, conMenu_Edit_NewPath, "新增路径(&A)"
                 .Add xtpControlButton, conMenu_Edit_ModifyPath, "修改路径(&Q)"
                 .Add xtpControlButton, conMenu_Edit_DelPath, "删除路径(&W)"
            End With
        End If
        
        rptStPath.SetFocus
        If Not objPopup Is Nothing Then objPopup.ShowPopup
    End If
    
End Sub

Private Sub rptStPath_SelectionChanged()
'功能：保存选择的路径ID,并根据ID加载标准路径流程以及表单

    mintFocus = FC_PathList
    If rptStPath.Rows.Count <> 0 Then
        If Not rptStPath.SelectedRows(0).GroupRow Then
            If mlngStPathID <> Val(rptStPath.SelectedRows(0).Record.Tag) And Val(rptStPath.SelectedRows(0).Record.Tag) <> 0 Then
                mlngStPathID = Val(rptStPath.SelectedRows(0).Record.Tag)
                Call LoadPathByID(mlngStPathID, True, 0)
            End If
        End If
    End If
End Sub

Private Sub rtfPathCourse_GotFocus()
    mintFocus = FC_Course
End Sub

Private Sub rtfPathCourse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBar

    mintFocus = FC_Course
    If Button = 2 Then
        Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
        With objPopup.Controls
            .Add xtpControlButton, conMenu_Edit_NewCourseItem, "新增段落(&Z)"
            .Add xtpControlButton, conMenu_Edit_ModifyCourseItem, "修改段落(&U)"
            .Add xtpControlButton, conMenu_Edit_DelCourseItem, "删除段落(&D)"
        End With
        
        rtfPathCourse.SetFocus
        objPopup.ShowPopup
    End If
End Sub

Private Sub tbcPathName_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能:根据选项卡选择加载内容
    If Me.Visible Then
        mlngStPathID = 0
        Call LoadStPathList(Item.Index)
        Call LoadPathByID(mlngStPathID, True, 0)
    End If
End Sub

Private Sub tbcStPath_GotFocus()
    mintFocus = FC_TBC
End Sub

Private Sub tbcStPath_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBar
    
    mintFocus = FC_TBC
    
    If Button = 2 Then
        Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
        With objPopup.Controls
            .Add xtpControlButton, conMenu_Edit_NewTable, "新增表单(&K)"
            .Add xtpControlButton, conMenu_Edit_ModifyTable, "修改表单(&M)"
            .Add xtpControlButton, conMenu_Edit_DelTable, "删除表单(&Y)"
        End With
        
        tbcStPath.SetFocus
        objPopup.ShowPopup
    End If
End Sub

Private Sub tbcStPath_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：选择表单时加载表单内容

    If Me.Visible Then
        Call LoadPathByID(mlngStPathID, False, Item.Index)
        mintFocus = FC_TBC
        picPathCourse.Visible = Item.Index = 0
        picPathTable.Visible = Item.Index <> 0
    End If
    
End Sub

Private Sub LoadStPathList(ByVal lngIndex As Long)
'功能：加载标准路径目录
'参数:lngIndex 0-西医参考,1-中医参考

    Dim objRecord     As ReportRecord
    Dim objItem       As ReportRecordItem
    Dim i As Long, strDept As String
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select a.Id, a.科室名称, a.编码, a.路径名称, a.版本说明,疾病编码,手术编码 " & vbNewLine & _
        "From   标准路径目录 A,标准路径病种 B where A.ID=B.标准路径ID And Nvl(a.类别,0)=[1] " & vbNewLine & _
        " order by 科室名称,编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngIndex)
    rptStPath.Records.DeleteAll
    For i = 0 To rsTemp.RecordCount - 1
        Set objRecord = rptStPath.Records.Add
        Set objItem = objRecord.AddItem(rsTemp!ID & "")
        Set objItem = objRecord.AddItem(rsTemp!科室名称 & "")
        Set objItem = objRecord.AddItem(rsTemp!编码 & "")
        Set objItem = objRecord.AddItem(rsTemp!路径名称 & "")
        Set objItem = objRecord.AddItem(rsTemp!版本说明 & "")
        Set objItem = objRecord.AddItem(rsTemp!疾病编码 & "")
        Set objItem = objRecord.AddItem(rsTemp!手术编码 & "")
        objRecord.Tag = CStr(rsTemp!ID)
        rsTemp.MoveNext
    Next
    rptStPath.Populate
    
    For i = 0 To rptStPath.Rows.Count - 1
        If mlngStPathID = 0 Then
            If rptStPath.Rows(i).GroupRow = False Then
                rptStPath.Rows(i).Selected = True
                mlngStPathID = Val(rptStPath.Rows(i).Record(COL_ID).Value)
                rptStPath.SelectedRows(0).ParentRow.Expanded = True
                Exit For
            End If
        Else
            If rptStPath.Rows(i).GroupRow = False Then
                If mlngStPathID = Val(rptStPath.Rows(i).Record(COL_ID).Value) Then
                    rptStPath.Rows(i).Selected = True
                    rptStPath.SelectedRows(0).ParentRow.Expanded = True
                    Exit For
                End If
            End If
        End If
    Next
    
    '不会自动调用LoadPathByID(mlngStPathID, True)
    Call LoadPathByID(mlngStPathID, True)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub LoadPathByID(ByVal lngId As Long, Optional ByVal blnReadData As Boolean, Optional ByVal lng序号 As Long)
'功能：根据选择的标准路径ID读取数据，并根据表单序号加载路径流程，路径表单，表单表头
'参数：lngID   选择的路径ID
'      blnReadData 是否读取标准路径信息（在标准路径初次加载或者标准路径切换时是均需要读取）
'      lng序号  0 标准路径流程，1 表单1，2，表单2...
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long, k As Long
    Dim strSql As String, strFilter As String
    Dim strTilePos As String '记录标题行位置格式为，标题1起始位置，长度;标题2起始位置，长度
    Dim lngColCount As Long, lng表单行数 As Long, lngBeginRow As Long
    Dim lngRowCount As Long
    Dim strContent As String
    Dim strB As String
    Dim arrTemp() As String
    Dim n As Integer, intPos As Integer
    
    On Error GoTo errH
    
    If blnReadData Then
        '删除选项卡，清空vs数据
        vsPathTable.Delete
        For i = tbcStPath.ItemCount - 1 To 1 Step -1
            tbcStPath.RemoveItem (i)
        Next
        
        '加载标准住院流程
        rtfPathCourse.Visible = False
        rtfPathCourse.Text = ""
        strSql = "Select 标题, 内容 From 标准路径流程 Where 标准路径id = [1] Order By 序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        mstrTilePos = ""
        strB = "β,γ,α,μ,δ,κ,ц,α,λ,℃,Δ"
        arrTemp = Split(strB, ",")
        If rsTmp.RecordCount <> 0 Then
            For i = 1 To rsTmp.RecordCount
                strTilePos = strTilePos & ";" & Len(strContent) & "," & Len(rsTmp!标题) & "," & n
                mstrTilePos = mstrTilePos & "," & Len(strContent)
                n = 0
                For j = LBound(arrTemp) To UBound(arrTemp)
                    Do
                        intPos = InStr(intPos + 1, "" & rsTmp!内容, arrTemp(j))
                        If intPos = 0 Then Exit Do
                        n = n + 1
                        DoEvents
                    Loop
                Next
                strContent = strContent & rsTmp!标题 & vbNewLine & vbNewLine & rsTmp!内容 & vbNewLine & vbNewLine
                rsTmp.MoveNext
            Next
            rtfPathCourse.Text = strContent
            mstrTilePos = Mid(mstrTilePos, 2)
        End If
               
        
        Call SetStPathCourceFont(Mid(strTilePos, 2)) '设置字体
        rtfPathCourse.Visible = True
        
        '读取表单总体信息
        strSql = "Select a.表单序号 表单序号, b.表单名称, b.表单表头, a.行数, a.列数" & vbNewLine & _
                "From (Select 表单序号, Max(分类序号) 行数, Max(阶段序号) 列数 From 标准路径表单 Where 标准路径id = [1] Group By 表单序号) A, 标准路径表单 B" & vbNewLine & _
                "Where b.标准路径id =[1] And a.表单序号 = b.表单序号 And b.分类序号 = 1 And b.阶段序号 = 1" & vbNewLine & _
                "Order By 表单序号"

        Set mrs表头信息 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        
        '加载标准路径表单选项
        If mrs表头信息.RecordCount > 0 Then
            j = mrs表头信息.RecordCount
            For i = 1 To j
                mrs表头信息.Filter = "表单序号 =" & i
                Call tbcStPath.InsertItem(i, mrs表头信息!表单名称, picPathTable.hwnd, 0)
            Next
            '读取表单数据
            strSql = "Select  表单序号, 表单名称, 表单表头, 分类序号, 分类名称, 阶段序号, 阶段名称, 路径内容" & vbNewLine & _
                "From   标准路径表单" & vbNewLine & _
                "where 标准路径id=[1]"
            Set mrs表单 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        End If
    End If
    
    If lng序号 <> 0 Then
        '没有表单信息数据则不加载表单信息
        
        mrs表单.Filter = ""
        If mrs表单.RecordCount = 0 Then tbcStPath.Item(0).Selected = True: Exit Sub
        mrs表头信息.Filter = " 表单序号 =" & lng序号
        If mrs表头信息.RecordCount = 0 Then tbcStPath.Item(0).Selected = True: Exit Sub
        
        '加载表单表头
        lblTableTile.Caption = ""
        lblTableTile.Caption = vbNewLine & mrs表头信息!表单表头
        
        With vsPathTable
            .Redraw = False
            .Rows = 0
            .Cols = 0
            '确定行数
            lngColCount = Val(mrs表头信息!列数 & "")
            '确定总行数
            lng表单行数 = Val(mrs表头信息!行数 & "")
            lngRowCount = IntEx(lngColCount / (M_INT_STEPNUM + 1)) * lng表单行数 + IntEx(lngColCount / (M_INT_STEPNUM + 1)) - 1
            If lngRowCount = 1 And lngColCount = 1 Then
                .Rows = 0
                .Cols = 0
                Call SetVsStyle
                Call picPathTable_Resize '由于lblTableTile是autoSize的因此需要调用resize
                tbcStPath.Item(lng序号).Selected = True
                Exit Sub
            Else
                .Rows = lngRowCount
                .Cols = IIf(lngColCount > M_INT_STEPNUM, M_INT_STEPNUM + 1, lngColCount)
            End If
    
    
            For k = 1 To IntEx(lngColCount / (M_INT_STEPNUM + 1))
                lngBeginRow = (k - 1) * lng表单行数 + (k - 1)
                For i = lngBeginRow To lngBeginRow + lng表单行数 - 1
                    For j = 0 To .Cols - 1
                        '每个表单表格区域的第一个单元格为时间
                        If i = lngBeginRow And j = 0 Then
                            .TextMatrix(i, j) = "时间"
                        Else
                            If Not (i = lngBeginRow Or j = 0) Then
                                strFilter = "表单序号=" & lng序号 & " and 分类序号=" & i - lngBeginRow + 1 & " and 阶段序号=" & (k - 1) * 3 + j + 1
                                mrs表单.Filter = strFilter
                                If mrs表单.RecordCount = 1 Then
                                    .TextMatrix(i, j) = Nvl(mrs表单!路径内容, " ")
                                    .TextMatrix(i, 0) = Replace(Replace(Replace(mrs表单!分类名称 & "", Chr(13), ""), Chr(10), ""), " ", "")
                                    .TextMatrix(lngBeginRow, j) = mrs表单!阶段名称 & ""
                                End If
                            End If
                        End If
                    Next
                Next
            Next
            
            Call SetVsStyle
            .Redraw = True
            Call picPathTable_Resize '由于lblTableTile是autoSize的因此需要调用resize
        
            
        End With
    End If
    
    tbcStPath.Item(lng序号).Selected = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStPathCourceFont(ByVal strTilePos As String)
'功能：对RichTextBox进行字体设置
'参数 strTilePos 记录标题行位置格式为，标题1起始位置，标题长度;标题2起始位置，标题长度
    Dim arrTmp As Variant, i As Long, j As Long
    
    On Error Resume Next
    If Len(Trim(strTilePos)) = 0 Then Exit Sub
    arrTmp = Split(Trim(strTilePos), ";")
    
    With rtfPathCourse
        For i = LBound(arrTmp) To UBound(arrTmp)
            If Val(Split(arrTmp(i), ",")(2)) = 0 Then
                .SelStart = Split(arrTmp(i), ",")(0)
            Else
                j = i
                Exit For
            End If
            .SelLength = Split(arrTmp(i), ",")(1)
            .SelFontSize = 14
            .SelFontName = "黑体"
            .SelBold = True
            .SelLength = 0
        Next
        For i = j To UBound(arrTmp)
            .SelStart = Val(Split(arrTmp(i), ",")(0)) - Val(Split(arrTmp(j), ",")(2))
            .SelLength = Split(arrTmp(i), ",")(1)
            .SelFontSize = 14
            .SelFontName = "黑体"
            .SelBold = True
            .SelLength = 0
        Next
        .SelStart = 0 '光标移动到开始
        
    End With
    
End Sub

Private Sub SetVsStyle()
'功能：根据内容设置表单表格的单元格的高度与宽度,以及内容颜色等，以及单元格的合并等

    Dim i As Long, j As Long
    Dim lngmaxHeight As Long
    
    
   On Error GoTo errH
    With vsPathTable
        If .Rows = 0 And .Cols = 0 Then Exit Sub
        '修改分类名称，阶段，分类加粗居中
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = 4 '居中
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = &HE1FFE1
        
        .AutoResize = False
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1, False, 0) '自动调整大小
        '设置阶段字体，颜色，对齐方式
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = "时间" Then
                .Cell(flexcpAlignment, i, 0, i, .Cols - 1) = 4
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False '设置加粗前要先清除加粗
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = True
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE1FFE1
            Else
                If .Cols > 1 Then
                    .Cell(flexcpAlignment, i, 1, i, .Cols - 1) = 0
                End If
            End If
        Next
        
        '获取同一行最高的单元格高度赋值给行高
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                For j = 0 To .Cols - 1
                    If j = 0 Then
                        lngmaxHeight = ComputerLines(.TextMatrix(i, j))
                    Else
                        lngmaxHeight = IIf(lngmaxHeight > ComputerLines(.TextMatrix(i, j)), lngmaxHeight, ComputerLines(.TextMatrix(i, j)))
                    End If
                Next
                .RowHeight(i) = IIf(lngmaxHeight = 0, 5, lngmaxHeight) * Me.TextHeight("字") * 1.5
            Else
                For j = 0 To .Cols - 1
                    .TextMatrix(i, j) = " " '为了合并单元格
                Next
            End If
        Next
        '分割行单元格合并，以及边框颜色设置
        .MergeCells = flexMergeFree
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = " " Then
                Call .CellBorderRange(i, 0, i, .Cols - 1, &HFFFFFF, 1, 0, 1, 0, 1, 0)
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HFFFFFF
                .MergeRow(i) = True
            End If
        Next
        
        For i = 1 To .Cols - 1
            .ColWidth(i) = 4000
        Next
        .ColWidth(0) = 1500
        '实现自由拖动列宽
        .FixedRows = 1
        Call .CellBorderRange(0, 0, 0, .Cols - 1, &H8000&, 0, 0, 1, 1, 1, 1)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function ComputerLines(ByVal strInput As String) As Long
'功能：计算输入文本中回车符的个数
'参数：  strInput   要计算回车符的字符串
'返回：   回车符的个数

    Dim strTmp As String
    Dim Count  As Long, lngPos As Long, lngLen As Long
    
    lngPos = InStr(strInput, Chr(13))
    lngLen = Len(strInput)
    strTmp = strInput
    
    Do While lngPos <> 0
        If Trim(strTmp) = "" Then Exit Do
        If lngPos + 1 <= lngLen Then
            strTmp = Mid(strTmp, lngPos + 1)
            Count = Count + 1
            lngPos = InStr(strTmp, Chr(13))
            lngLen = Len(strTmp)
        End If
    Loop
    
    ComputerLines = Count + 2
    
End Function

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_ImportPathTable, "导入路径表单")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewPath, "新增路径(&A)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyPath, "修改路径(&Q)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DelPath, "删除路径(&W)")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewCourseItem, "新增段落(&Z)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyCourseItem, "修改段落(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DelCourseItem, "删除段落(&D)")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewTable, "新增表单(&K)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyTable, "修改表单(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DelTable, "删除表单(&Y)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyTableContent, "修改内容(&G)")
        objControl.BeginGroup = True
    End With

   Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
                objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
                objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
        End With
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "展开/折叠组(&X)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&I)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一个(&N)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True '固有
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With
    
    '查找项特殊处理
    '-----------------------------------------------------
    '主菜单右侧的查找
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, 0, "查找")
        objControl.IconId = conMenu_View_Find
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Caption = ""
        objCustom.Handle = txtFind.hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewPath, "新增路径")
            objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
            
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewCourseItem, "新增段落")
            objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyCourseItem, "修改段落")
            objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DelCourseItem, "删除段落")
            objControl.Style = xtpButtonIconAndCaption
            
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewTable, "新增表单")
            objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyTable, "修改表单")
            objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DelTable, "删除表单")
            objControl.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyTableContent, "修改内容")
            objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
            
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    

    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewPath '新增路径
        .Add FCONTROL, vbKeyM, conMenu_Edit_DelPath '修改路径
        .Add FSHIFT, vbKeyD, conMenu_Edit_DelPath '删除路径
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找
        .Add 0, vbKeyF3, conMenu_View_FindNext '查找下一个
        
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '折叠所有组
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
End Sub

Private Sub InitPathList()
'功能：初始化路径列表
    Dim objCol        As ReportColumn
    
    With rptStPath
        '初始化Report控件的列与属性
        Set objCol = .Columns.Add(PathListCols.COL_ID, "ID", 20, False)
            objCol.Alignment = xtpAlignmentCenter: objCol.Resizable = True: objCol.AllowDrag = False: objCol.Visible = False
        Set objCol = .Columns.Add(PathListCols.COL_科室名称, "科室名称", 80, False)
            objCol.Resizable = True: objCol.Alignment = xtpAlignmentLeft: objCol.AllowDrag = False: objCol.TreeColumn = True: objCol.Groupable = True
        Set objCol = .Columns.Add(PathListCols.COL_编码, "编码", 50, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_路径名称, "路径名称", 200, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_版本说明, "版本说明", 70, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_疾病编码, "疾病编码", 200, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_手术编码, "手术编码", 200, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoItemsText = "没有可显示的段落."
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '会引发SelectionChanged事件
        
        .GroupsOrder.Add rptStPath.Columns(COL_科室名称)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add rptStPath.Columns(COL_科室名称)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add rptStPath.Columns(COL_路径名称)
        .SortOrder(1).SortAscending = True
    '定位到选择的标准路径
    End With
End Sub

Private Function DeleteItem(ByVal lngDelType As Long, Optional ByVal lngStPathID As Long)
'功能：删除特定段落
'lngDelType 0-删除标准路径
'           1-删除路径流程段落
'           2-删除标准路径表单
'lngStPathID 标准路径ID
    Dim strSql As String
    
    On Error GoTo errH
    
    Select Case lngDelType
        Case 0
            If MsgBox("你确定要删除" & rptStPath.SelectedRows(0).Record(COL_路径名称).Value & "吗？", vbYesNo, gstrSysName) = vbYes Then
                strSql = "Zl_标准路径目录_Delete(" & mlngStPathID & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                If rptStPath.SelectedRows(0).ParentRow.Childs.Count > 1 Then
                    If Val(rptStPath.SelectedRows(0).ParentRow.Childs(0).Record(COL_ID).Value) <> mlngStPathID Then
                        mlngStPathID = Val(rptStPath.SelectedRows(0).ParentRow.Childs(0).Record(COL_ID).Value)
                    Else
                        mlngStPathID = Val(rptStPath.SelectedRows(1).ParentRow.Childs(0).Record(COL_ID).Value)
                    End If
                Else
                    mlngStPathID = 0
                End If
                Call LoadStPathList(tbcPathName.Selected.Index)
            End If
        Case 1
            If frmStPathItemEdit.ShowMe(Me, 2, mlngStPathID, GetCourseItemNo(rtfPathCourse.SelStart)) = True Then
                Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
            End If
        Case 2
            If MsgBox("你确定要删除" & tbcStPath.Selected.Caption & "吗？", vbYesNo, gstrSysName) = vbYes Then
                strSql = "Zl_标准路径表单_Delete(" & mlngStPathID & "," & tbcStPath.Selected.Index & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                Call LoadPathByID(mlngStPathID, True, 0)
            End If
    End Select
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ModItem(ByVal lngDelType As Long, Optional ByVal lngStPathID As Long)
'功能：修改特定段落
'lngDelType 0-修改标准路径
'           1-修改路径流程段落
'           2-修改标准路径表单
'lngStPathID 标准路径ID
    Dim strSql As String
    
    Select Case lngDelType
        Case 0
            With rptStPath.SelectedRows(0)
                If frmStPathEdit.ShowMe(Me, 1, Val(.Record(COL_ID).Value), .Record(COL_路径名称).Value, .Record(COL_编码).Value, _
                    .Record(COL_科室名称).Value, .Record(COL_版本说明).Value, .Record(COL_疾病编码).Value, .Record(COL_手术编码).Value, tbcPathName.Selected.Index) = True Then
                    Call LoadStPathList(tbcPathName.Selected.Index)
                End If
            End With
        Case 1
            If frmStPathItemEdit.ShowMe(Me, 1, mlngStPathID, GetCourseItemNo(rtfPathCourse.SelStart)) = True Then
                Call LoadPathByID(mlngStPathID, False, tbcStPath.Selected.Index)
            End If
        Case 2
            If frmStTableEdit.ShowMe(Me, mlngStPathID, tbcStPath.Selected.Index, tbcStPath.Selected.Caption, lblTableTile.Caption) = True Then
                Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
            End If
        Case 3
            If frmStTableContent.ShowMe(Me, mlngStPathID, tbcStPath.Selected.Index) = True Then
                Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
            End If
    End Select

End Function

Private Function InsertItem(ByVal lngDelType As Long)
'功能：插入特定段落
'lngDelType 0-插入标准路径
'           1-插入路径流程段落
'           2-插入标准路径表单
'lngStPathID 标准路径ID
    Dim strSql As String, strDep As String
    
    Select Case lngDelType
        Case 0
            If rptStPath.Rows.Count > 0 Then
                If Not rptStPath.SelectedRows(0).GroupRow Then
                    strDep = rptStPath.SelectedRows(0).Record(COL_科室名称).Value
                Else
                    strDep = rptStPath.SelectedRows(0).Childs(0).Record(COL_科室名称).Value
                End If
            End If
            
            If frmStPathEdit.ShowMe(Me, 0, mlngStPathID, , , strDep, , , , tbcPathName.Selected.Index) = True Then
                Call LoadStPathList(tbcPathName.Selected.Index)
            End If
        Case 1
            If tbcStPath.Selected.Index <> 0 Then tbcStPath.Item(0).Selected = True
            If frmStPathItemEdit.ShowMe(Me, 0, mlngStPathID, GetCourseItemNo(rtfPathCourse.SelStart, True)) = True Then
                Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
            End If
        Case 2
            If frmStTableEdit.ShowMe(Me, mlngStPathID) = True Then
                tbcStPath.Item(tbcStPath.ItemCount - 1).Selected = True
                Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
            End If
    End Select

End Function

Private Sub txtFind_GotFocus()
'功能：获得焦点，全选
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    
End Sub

Private Sub vsPathTable_DblClick()
'功能：进入标准路径表单编辑界面
    If frmStTableContent.ShowMe(Me, mlngStPathID, tbcStPath.Selected.Index) = True Then
        Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
    End If
End Sub

Private Sub vsPathTable_GotFocus()
    mintFocus = FC_Table
End Sub

Private Sub vsPathTable_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBar

    If Button = 2 Then
        Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
        With objPopup.Controls
            .Add xtpControlButton, conMenu_Edit_ModifyTableContent, "修改内容(&G)"
        End With
        vsPathTable.SetFocus
        objPopup.ShowPopup
    End If
End Sub

Private Function GetCourseItemNo(ByVal lngSelStar As Long, Optional ByVal blnNew As Boolean = False) As Long
'功能：获得当前光标所对应的路径流程段落序号
'参数：lngSelStar 当前光标位置
'      blnNew    是否是新增操作
'返回： 长整型   当前段落序号
    Dim arrTmp As Variant, i As Long, lngNo As Long
    
    '没有定义路径流程段落
    If Len(Trim(mstrTilePos)) = 0 Then GetCourseItemNo = 1: Exit Function
    '光标在开始位置
    If lngSelStar = 0 Then GetCourseItemNo = 1: Exit Function
    
    arrTmp = Split(Trim(mstrTilePos), ",")
    '仅有一个流程段落时的处理
    If LBound(arrTmp) = UBound(arrTmp) Then GetCourseItemNo = IIf(blnNew, 2, 1): Exit Function
    '其他情况的处理
    lngNo = 0
    For i = LBound(arrTmp) To UBound(arrTmp)
        If i < UBound(arrTmp) Then
            If lngSelStar >= Val(arrTmp(i)) And lngSelStar < Val(arrTmp(i + 1)) Then
                lngNo = IIf(blnNew, i + 2, i + 1): Exit For
            End If
        End If
    Next
    If lngNo = 0 Then lngNo = IIf(blnNew, UBound(arrTmp) + 2, UBound(arrTmp) + 1)
    
    GetCourseItemNo = lngNo
    
End Function

Private Sub FindPath(Optional ByVal blnNext As Boolean)
'功能：查找(下一条）数据
'参数：blnNext=是否查找下一条
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    Dim strInput As String
    Dim intType As Integer
    Dim lngRow As Long
    
    
    If Trim(txtFind.Text) = "" Then Exit Sub
     
    '开始查找行
    With rptStPath
        If .SelectedRows.Count > 0 Then
            If Not .SelectedRows(0).GroupRow Then blnHave = True: lngRow = .SelectedRows(0).Index
        End If
        If Not blnNext Or blnReStart Or Not blnHave Then
            i = 0 'ReportControl的索引从是0开始
        Else
            i = .SelectedRows(0).Index + 1
        End If

        '查找路径
        For i = i To .Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If .Rows(i).Record(COL_路径名称).Value Like "*" & Trim(txtFind.Text) & "*" Then Exit For
            End If
        Next
        
        If i <= .Rows.Count - 1 Then
            blnReStart = False
            '该行选中且显示在可见区域,并引发SelectionChanged事件
            .SetFocus
            Set .FocusedRow = .Rows(i)
            .Rows(i).ParentRow.Expanded = True
        Else
            blnReStart = True
            MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的标准路径。", vbInformation, gstrSysName
        End If
    
    End With
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
'功能:记录表打印
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objReport As ReportControl
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    
    Set objReport = rptStPath
    strSubhead = "标准路径清单"
    
    If objReport.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlControl.RPTCopyToVSF(objReport, vsTemp) Is Nothing Then Exit Sub
    
    '调用打印部件处理
    
    Set objPrint.Body = Me.vsTemp
    objPrint.Title.Text = strSubhead
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人:" & UserInfo.姓名)
    Call objAppRow.Add("打印时间:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub



