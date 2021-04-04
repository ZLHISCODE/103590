VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport1 
   AutoRedraw      =   -1  'True
   Caption         =   "单病种相关非特异性指标评估表"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   Icon            =   "frmReport1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11835
   Begin VB.ComboBox cboDate 
      Height          =   300
      Left            =   9720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1320
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8130
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmReport1.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17965
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
   Begin VB.Frame fraReport 
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   11655
      Begin VSFlex8Ctl.VSFlexGrid vsItem 
         Height          =   7530
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   11535
         _cx             =   20346
         _cy             =   13282
         Appearance      =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14744288
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   350
         ColWidthMin     =   0
         ColWidthMax     =   8000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReport1.frx":70E6
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmReport1.frx":71A6
      Left            =   825
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String

Private mstr路径名称 As String
Private mlng路径ID  As Long
Private mblnEdit As Boolean '当前是否在编辑模式
Private mstr当前期间 As String '当前选择的期间
Private mstr前一期间 As String
Private mrsDate As ADODB.Recordset '期间
Private mrsPati As ADODB.Recordset '某个病种在指定出院时间范围内的病人

Private Const conMenu_Date = 400
Private Const conMenu_Edit_GetCur = 3052
Private Const conMenu_Edit_GetAll = 3013

Private Enum 列名
    COL序号 = 0
    COL项目文本1 = 1
    COL项目文本2 = 2
    col结果 = 3
    col备注 = 4
End Enum


Public Function ShowMe(frmMain As Object, ByVal lng路径ID As Long, ByVal str路径名称 As String) As Boolean
    mstr路径名称 = str路径名称
    mlng路径ID = lng路径ID
    
    Me.Show 1, frmMain
End Function

Private Sub cboDate_Click()
    
    If cboDate.ListIndex >= 0 Then
        mrsDate.Filter = "ID=" & cboDate.ItemData(cboDate.ListIndex)
        Me.dkpMain.Panes(1).Title = "时间范围：" & Format(mrsDate!开始时间, "yyyy年mm月dd日") & _
                    "到" & Format(mrsDate!结束时间, "yyyy年mm月dd日") & "        病种：" & mstr路径名称
        mstr前一期间 = mstr当前期间
        mstr当前期间 = mrsDate!期间
        Call zlRefresh
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And mblnEdit Then
        Call ExeCancelEdit
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    mblnEdit = False
    fraReport.Visible = False
    mstrPrivs = gstrPrivs
    
    Call InitCommandBar
    Call InitDockPannel
      
    Call InitListTable
    
    Call FillStructure
    Call FillDate '填充报表数据
    
    Call RestoreWinState(Me, App.ProductName)
        
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    mstr当前期间 = ""
    mstr前一期间 = ""
    Set mrsDate = Nothing
    Set mrsPati = Nothing
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustomControl As CommandBarControlCustom

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '菜单定义:包括公共部份
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_GetAll, "读取所有数据(&R)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_GetCur, "读取当前行数据(&X)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend, "编辑SQL(&E)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
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
    
    '主菜单右侧的查找
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, 0, "期间 ")
        objControl.Flags = xtpFlagRightAlign
        Set objCustomControl = .Add(xtpControlCustom, conMenu_Date, "")
        objCustomControl.Handle = cboDate.Hwnd
        objCustomControl.Flags = xtpFlagRightAlign
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_GetAll, "读取所有数据")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_GetCur, "读取当前行数据")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend, "编辑SQL")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
                        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each objControl In objBar.Controls
        If objControl.Type = xtpControlButton Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新增
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改
        
        .Add 0, vbKeyF6, conMenu_Edit_GetCur '读取当前行
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF4, conMenu_Edit_Compend '编辑SQL
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save '保存
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save '保存
        
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
End Sub

Private Sub InitDockPannel()
    Dim objPane As Pane

    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True

    Set objPane = Me.dkpMain.CreatePane(1, 600, 600, DockTopOf, Nothing)
    objPane.Title = "病种：" & mstr路径名称
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    objPane.MinTrackSize.SetSize Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub InitListTable()
'功能：初始化单据清单表格
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "序号,1000,4;评估指标,1800,4;评估指标,3200,4;结果,1200,7;备注,2000,7"
    arrHead = Split(strHead, ";")
    With vsItem
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
    End With
End Sub
'
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = fraReport.Hwnd
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
'    '窗体其它控件Resize处理
    vsItem.Width = fraReport.Width
    vsItem.Height = fraReport.Height
End Sub


Private Sub FuncPathTableOutput(bytStyle As Byte)
'功能：输出报表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim rsTmp As ADODB.Recordset
    
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim lngRow As Long, lngCol As Long
    Dim lngColor As Long, bytR As Byte
        
    '表头
    objOut.Title.Text = Me.Caption
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    Set objRow = New zlTabAppRow
    objRow.Add " "
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "                                              病种：" & mstr路径名称
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    If mstr当前期间 <> "" Then
        objRow.Add "时间范围：" & Format(mrsDate!开始时间, "yyyy年mm月dd日") & _
                        "到" & Format(mrsDate!结束时间, "yyyy年mm月dd日")
    End If
    objOut.UnderAppRows.Add objRow
    
    '表下
'    Set objRow = New zlTabAppRow
'    objRow.Add "打印人：" & UserInfo.姓名
'    objRow.Add "打印日期：" & Format(zldatabase.Currentdate(), "yyyy年MM月dd日")
'    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = vsItem
    
    '输出
    With vsItem
        If bytStyle = 1 Then
            bytR = zlPrintAsk(objOut)
            Me.Refresh
            If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
        Else
            zlPrintOrView1Grd objOut, bytStyle
        End If
    End With
   
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    
    Select Case Control.ID
     '0.输出
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_Print
        Call FuncPathTableOutput(1)
    Case conMenu_File_Preview
        Call FuncPathTableOutput(2)
    Case conMenu_File_Excel
        Call FuncPathTableOutput(3)
            
    Case conMenu_Edit_NewItem '新增
        Call ExeNew
        
    Case conMenu_Edit_Modify '修改
        Call ExeModify
        
    Case conMenu_Edit_Delete '删除
        Call ExeDelete
   
    Case conMenu_Edit_GetAll    '提取所有数据
        Call ExeGetData
    Case conMenu_Edit_GetCur    '提取当前行数据
        If vsItem.Row < vsItem.FixedRows Then
            Me.stbThis.Panels(2) = "请先选中报表中的一行。"
            Exit Sub
        End If
        Call ExeGetData(vsItem.Row)
    Case conMenu_Edit_Compend   '编辑SQL
        If vsItem.Row < vsItem.FixedRows Then
            Me.stbThis.Panels(2) = "请先选中报表中的一行。"
            Exit Sub
        End If
        
        Call ExeDefineSQL
        
    Case conMenu_Edit_Transf_Save   '保存
        Call ExeSaveData
    Case conMenu_Edit_Transf_Cancle '取消
        Call ExeCancelEdit
   
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = _
                IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
  
    Case conMenu_View_Refresh '刷新
        Call FillStructure
        Call zlRefresh
        Me.stbThis.Panels(2).Text = "操作成功。"
   
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.Hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.Hwnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '退出
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    If mblnEdit Then    '编辑状态
        Select Case Control.ID
            Case conMenu_EditPopup, conMenu_HelpPopup, conMenu_Edit_GetAll, conMenu_Edit_GetCur, conMenu_Edit_Transf_Save, _
                conMenu_Edit_Transf_Cancle, conMenu_Edit_Compend  '提取数据,保存,取消,定义数据源
                Control.Enabled = True
            Case Else
                Control.Enabled = False
        End Select
    Else
        
        '根据权限设置按钮可见状态
        Call SetControlVisible(Control)
        If Not Control.Visible Then Exit Sub
    
        Select Case Control.ID
            Case conMenu_Edit_Modify, conMenu_Edit_Delete
                Control.Enabled = mstr当前期间 <> ""
                
            Case conMenu_Edit_GetAll, conMenu_Edit_GetCur, conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle '提取数据,保存,取消
                Control.Enabled = False
                 
            Case conMenu_View_ToolBar_Button '工具栏
                If cbsMain.count >= 2 Then
                    Control.Checked = Me.cbsMain(2).Visible
                End If
            Case conMenu_View_ToolBar_Text '图标文字
                If cbsMain.count >= 2 Then
                    Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
                End If
            Case conMenu_View_ToolBar_Size '大图标
                Control.Checked = Me.cbsMain.Options.LargeIcons
            Case conMenu_View_StatusBar '状态栏
                Control.Checked = Me.stbThis.Visible
            Case Else
                Control.Enabled = True
        End Select
    End If
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置菜单和工具栏的可见状态
    Dim blnVisible As Boolean, strItem As String

    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Category = "已判断" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Compend
            If InStr(";" & mstrPrivs & ";", ";单病种报表定义;") = 0 Then blnVisible = False
    End Select
    
    Control.Visible = blnVisible
    Control.Category = "已判断"
End Sub

Private Sub SetVsitemEdit(ByVal blnEnable As Boolean)
'功能：设置表格的可输入状态
'参数：bytEnable:=0不允许输入,1=允许输入
    
    If blnEnable Then
        vsItem.Editable = flexEDKbdMouse
        mblnEdit = True
    Else
        vsItem.Editable = flexEDNone
        mblnEdit = False
    End If
End Sub

Private Sub ExeCancelEdit()
    
    If MsgBox("放弃操作后，本次编辑的内容将不会被保存，你确定要继续吗？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
        Exit Sub
    End If
    Call SetVsitemEdit(False)
    Call zlRefresh
End Sub

Private Sub ExeSaveData()
'功能：保存数据
    Dim lng文件ID As Long, strContent As String
    Dim i As Long, strSQL As String
    Dim blnAllNull As Boolean
    
    On Error GoTo errH
    lng文件ID = cboDate.ItemData(cboDate.ListIndex)
    
    With vsItem
        For i = .FixedRows To .Rows - 1
            '行号|结果|备注||...,末尾带||,结果或备注为空是要留一个空格
            strContent = strContent & .RowData(i) & "|" & IIf(Trim(.TextMatrix(i, col结果)) = "", " ", Trim(.TextMatrix(i, col结果))) & "|" & IIf(Trim(.TextMatrix(i, col备注)) = "", " ", Trim(.TextMatrix(i, col备注))) & "||"
            
        Next
        If blnAllNull Then strContent = ""
    End With
    
    strSQL = "Zl_路径报表文件_Update(" & lng文件ID & ",'" & strContent & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    stbThis.Panels(2).Text = "数据保存成功。"
    Call SetVsitemEdit(False)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExeDelete()
'功能：执行数据删除操作
    Dim lngId As Long, strSQL As String
        
    If MsgBox("你确定要删除" & cboDate.Text & "的报表吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    On Error GoTo errH
    
    lngId = cboDate.ItemData(cboDate.ListIndex)
    
    strSQL = "Zl_路径报表文件_Delete(" & lngId & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    stbThis.Panels(2).Text = cboDate.Text & "的报表删除成功。"
    
    Call FillDate
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExeModify()
'功能：执行修改数据操作
    Call SetVsitemEdit(True)

End Sub

Private Sub ExeNew()
'功能：新增报表
    
    frmReport1Add.mlng路径ID = mlng路径ID
    frmReport1Add.Show vbModal, gfrmMain
    
    If frmReport1Add.mblnOK Then
        Call SetVsitemEdit(True)
        Call FillDate(frmReport1Add.mstr期间)
    End If
End Sub

Private Sub FillDate(Optional ByVal str期间 As String)
'功能：加载当前路径表的期间
'      str期间:缺省的当前期间
    Dim strSQL As String
    Dim i As Long
 
    strSQL = "Select 期间,ID,开始时间,结束时间 From 路径报表文件 Where 报表ID = 1 And 路径ID = [1] Order by 期间 Desc"
    On Error GoTo errH
    Set mrsDate = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng路径ID)
        
    cboDate.Clear
    For i = 1 To mrsDate.RecordCount
        cboDate.AddItem Mid(mrsDate!期间, 1, 4) & "年" & Mid(mrsDate!期间, 5, 2) & "月"
        cboDate.ItemData(cboDate.NewIndex) = mrsDate!ID
        mrsDate.MoveNext
    Next
    If cboDate.ListCount > 0 Then
        If str期间 = "" Then
            cboDate.ListIndex = 0
        Else
            cbo.Locate cboDate, Mid(str期间, 1, 4) & "年" & Mid(str期间, 5, 2) & "月"
        End If
    Else
        dkpMain.Panes(1).Title = "病种：" & mstr路径名称
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub FillStructure()
'功能：刷新报表结构
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
 
    strSQL = "Select 行号, 项目序号, 项目文本1, 项目文本2, Sql文本 From 路径报表结构 Where 报表id = 1 Order By 行号"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        With vsItem
            .Rows = .FixedRows
            .Rows = .FixedRows + rsTmp.RecordCount
            
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val("" & rsTmp!行号)
                .TextMatrix(i, COL序号) = "" & rsTmp!项目序号
                .TextMatrix(i, COL项目文本1) = "" & rsTmp!项目文本1
                .TextMatrix(i, COL项目文本2) = NVL(rsTmp!项目文本2, rsTmp!项目文本1)
                .Cell(flexcpData, i, col结果) = "" & rsTmp!SQL文本
                .MergeRow(i) = True
                
                If Not IsNull(rsTmp!SQL文本) Then
                    .Cell(flexcpBackColor, i, col结果) = &HFFEFDF
                End If
                
                rsTmp.MoveNext
            Next
            
            .MergeCol(COL序号) = True
            .MergeCol(COL项目文本1) = True
            .MergeCol(col结果) = False
            .MergeCol(col备注) = False
        End With
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function zlRefresh() As Long
'功能：刷新报表数据
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, lngId As Long
    
    If cboDate.ListIndex < 0 Then
        With vsItem
            .Redraw = flexRDNone
            For i = 1 To .Rows - 1
                .TextMatrix(i, col结果) = ""
                .TextMatrix(i, col备注) = ""
            Next
            .Redraw = flexRDDirect
        End With
    Else
        lngId = cboDate.ItemData(cboDate.ListIndex)
        
        strSQL = "Select a.行号, a.项目值, a.备注" & vbNewLine & _
                "From 路径报表记录 A" & vbNewLine & _
                "Where a.文件id = [1]" & vbNewLine & _
                "Order By 行号"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngId)
        If rsTmp.RecordCount > 0 Then
            With vsItem
                .Redraw = flexRDNone
                For i = 1 To .Rows - 1
                    rsTmp.Filter = "行号=" & .RowData(i)
                    If rsTmp.RecordCount > 0 Then
                        .TextMatrix(i, col结果) = "" & rsTmp!项目值
                        .TextMatrix(i, col备注) = "" & rsTmp!备注
                    Else
                        .TextMatrix(i, col结果) = ""
                        .TextMatrix(i, col备注) = ""
                    End If
                Next
                .Redraw = flexRDDirect
            End With
        End If
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = col结果 Or Col = col备注) Then
        Cancel = True
    End If
End Sub

Private Sub ExeDefineSQL()
    Dim lng行号 As Long, strSQL文本 As String
    
    lng行号 = Val(vsItem.RowData(vsItem.Row))
    
    If lng行号 <> 0 Then
        '1,4,17行是特殊行，不能定义SQL
        If lng行号 = 1 Or lng行号 = 4 Or lng行号 = 17 Then
            MsgBox "当前行是标题行，不能定义SQL。", vbInformation, gstrSysName
            Exit Sub
        End If
    
        strSQL文本 = vsItem.Cell(flexcpData, vsItem.Row, col结果)
        Call frmReportSQLSet.ShowMe(gfrmMain, lng行号, strSQL文本)
        vsItem.Cell(flexcpData, vsItem.Row, col结果) = strSQL文本
    End If
End Sub

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With vsItem
            If (.Col = col结果 Or .Col = col备注) And mblnEdit Then .TextMatrix(.Row, .Col) = ""
        End With
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call GoNextCell
    End If
End Sub

Private Sub GoNextCell()
    If vsItem.Row < vsItem.Rows - 1 Then
        vsItem.Row = vsItem.Row + 1
        If vsItem.Col = vsItem.Cols - 1 Then vsItem.Col = col结果
    End If
End Sub

Private Sub vsItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call GoNextCell
    End If
End Sub

Private Sub ExeGetData(Optional ByVal lngRow As Long)
'功能：根据预定的SQL读取数据填充单元格
    Dim strSQL As String, strPati As String, strPatiPage As String
    Dim rsTmp As ADODB.Recordset
    Dim DatBegin As Date, DatEnd As Date, i As Long, lngS As Long, lngE As Long, lngErrCnt As Long
    Dim lng总住院天数 As Long, lng总费用 As Long, lng总人次 As Long, lng总人数 As Long
    Dim blnGeted总费用 As Boolean   '是否取过总费用
    Dim varPati As Variant, varPage As Variant
    Dim strParTable As String, strTablePati As String, strTablePage As String
    Dim strTempSQL As String
    Dim intMaxPati As Integer, intMaxPage As Integer

    On Error GoTo errH
    '读取病人信息
    If mrsPati Is Nothing Or mstr前一期间 <> mstr当前期间 Then
        DatBegin = mrsDate!开始时间
        DatEnd = mrsDate!结束时间
        strSQL = "Select a.病人id, a.主页id, a.疾病id" & vbNewLine & _
                "From (Select Row_Number() Over(Partition By a.病人id Order By a.主页id Desc, Decode(a.记录来源,4,1,a.记录来源) Desc,Sign(a.诊断类型-10), Decode(a.诊断类型,3,0,13,10,a.诊断类型) Desc) As Top, a.病人id, a.主页id," & vbNewLine & _
                "              a.疾病id" & vbNewLine & _
                "       From 病人诊断记录 A, 病案主页 B" & vbNewLine & _
                "       Where a.病人id = b.病人id And a.主页id = b.主页id And a.诊断次序 = 1 And a.诊断类型 In (1, 2, 3, 11, 12, 13) And" & vbNewLine & _
                "             b.出院日期 Between [2] And [3] And a.疾病id is Not Null) A, 临床路径病种 B" & vbNewLine & _
                "Where Top = 1 And a.疾病id = b.疾病id And b.性质=0 And b.路径ID = [1]"
        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng路径ID, DatBegin, DatEnd)
    End If
    
    If mrsPati.RecordCount > 0 Then
        mrsPati.MoveFirst
        For i = 1 To mrsPati.RecordCount
            strPatiPage = strPatiPage & "," & mrsPati!病人ID & ":" & mrsPati!主页ID
            If InStr(strPati & ",", "," & mrsPati!病人ID & ",") = 0 Then
                strPati = strPati & "," & mrsPati!病人ID
            End If
            mrsPati.MoveNext
        Next
        strPatiPage = Mid(strPatiPage, 2)
        strPati = Mid(strPati, 2)
        
        If Len(strPatiPage) > 4000 Then
            varPage = FuncGetTable(strPatiPage, 1, strTablePage, intMaxPage)
        End If
        If Len(strPati) > 4000 Then
            varPati = FuncGetTable(strPati, 0, strTablePati, intMaxPati)
        End If
    End If
    If strPatiPage = "" Then
        MsgBox "当前期间没有符合病种[" & mstr路径名称 & "]条件的病人。", vbInformation, gstrSysName
        Exit Sub
    End If
    lng总人次 = UBound(Split(strPatiPage, ",")) + 1
    lng总人数 = UBound(Split(strPati, ",")) + 1
    
    If lngRow = 0 Then
        lngS = vsItem.FixedRows
        lngE = vsItem.Rows - 1
    Else
        lngS = lngRow
        lngE = lngRow
    End If
    
    On Error Resume Next '如果出错，记录来下继续进行
    With vsItem
        For i = lngS To lngE
            stbThis.Panels(2).Text = "正在读取第" & i & "行数据"
            Me.Refresh
            .Cell(flexcpData, i, col备注) = ""  '清除前一次的
            .Cell(flexcpBackColor, i, col备注) = vbWhite    '取数出错后，备注列底色为浅黄色
                    
            strSQL = .Cell(flexcpData, i, col结果)
            If strSQL <> "" Then
                .TextMatrix(i, col结果) = ""
                '1:一、效率指标
                
                Select Case .RowData(i) '报表行号
                Case 2 '返回平均住院天数,总住院天数用于计算"日均费用"
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng路径ID)
                    Else
                        '参数号调整：将未超长的参数排在最前面
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '整体向后移动一位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then
                            .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                            If rsTmp.Fields.count > 1 Then lng总住院天数 = Val("" & rsTmp.Fields(1).Value)
                        End If
                    End If
                Case 3  '术前平均住院日
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '整体向后移动一位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                    
                '4:二、效果指标
                Case 5  '死亡率,治愈率,好转率
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总人次, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总人次, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    .TextMatrix(i + 1, col结果) = "": .TextMatrix(i + 2, col结果) = ""
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then
                            .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                            If rsTmp.Fields.count > 1 Then .TextMatrix(i + 1, col结果) = "" & rsTmp.Fields(1).Value
                            If rsTmp.Fields.count > 2 Then .TextMatrix(i + 2, col结果) = "" & rsTmp.Fields(2).Value
                        End If
                    End If
                Case 6, 7   '治愈率,好转率(如果定义了，缺省未定义)
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '整体向后移动一位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                Case 8  '院内感染数
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总人次, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总人次, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                Case 9  '手术部位感染率(如果定义了，缺省未定义)
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总人次, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总人次, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                Case 10 '14日再住院率,31日再住院率
                    .TextMatrix(i + 1, col结果) = ""
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPati, strPatiPage, lng总人次, mlng路径ID)
                    ElseIf Len(strPati) <= 4000 Then
                        Call FuncReDoSQLNum(strSQL, 3, 4)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 3) '整体向后移动3位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([4]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPati, lng总人次, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    Else
                        Call FuncReDoSQLNum(strSQL, 3, 4)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 3) '整体向后移动3位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([4]))"), strTempSQL)
                        
                        Call FuncReDoSQLNum(strSQL, 2, 13) '按照page最大取到10来推算Pati参数位置
                        strTempSQL = strTablePati
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPati, 12)  '整体向后移动12位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list([13]))"), strTempSQL)
                        
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总人次, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)), CStr(varPati(0)), CStr(varPati(1)), _
                            CStr(varPati(2)), CStr(varPati(3)), CStr(varPati(4)), CStr(varPati(5)), CStr(varPati(6)), CStr(varPati(7)), CStr(varPati(8)), CStr(varPati(9)))
    
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                        If rsTmp.Fields.count > 1 Then .TextMatrix(i + 1, col结果) = "" & rsTmp.Fields(1).Value
                    End If
                Case 11, 12 '31日再住院率,非计划重返手术室发生率(如果定义了，缺省未定义)
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总人次, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总人次, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
               
                    
                '-----------------------------------------------
                Case 13 '并发症发生率，及前三位疾病
                    .TextMatrix(i + 1, col结果) = "": .TextMatrix(i + 2, col结果) = "": .TextMatrix(i + 3, col结果) = ""
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总人次, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总人次, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then
                            .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                            If rsTmp.Fields.count > 1 Then .TextMatrix(i + 1, col结果) = "" & rsTmp.Fields(1).Value
                            If rsTmp.Fields.count > 2 Then .TextMatrix(i + 2, col结果) = "" & rsTmp.Fields(2).Value
                            If rsTmp.Fields.count > 3 Then .TextMatrix(i + 3, col结果) = "" & rsTmp.Fields(3).Value
                        End If
                    End If
                Case 14, 15, 16 '(如果定义了，缺省未定义)
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '整体向后移动一位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If

                '-----------------------------------------------------
                '17:三、工作量指标
                Case 18 '住院患者总人数(如果定义了，缺省未定义)
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, DatBegin, DatEnd, mlng路径ID)
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 19 '进入路径的患者总人次数
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '整体向后移动一位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 20 '完成路径的人次数
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '整体向后移动一位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                    
                Case 21 '变异病人数
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '整体向后移动一位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If

                '------------------------------------------------
                Case 23 '使用三代抗菌药物的患者比例
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总人次, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总人次, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 24 '使用抗生素的平均天数
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 2)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '整体向后移动一位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 25 '单病种次均费用,总费用(第2行取总住院天数) '如果只求一个单元格，则lng总住院天数没有值，需要SQL中自己取
                    blnGeted总费用 = True
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总人次, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总人次, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                        If rsTmp.Fields.count > 1 Then lng总费用 = Val("" & rsTmp.Fields(1).Value)
                    End If
                
                Case 26 '单病种日均费用(如果定义了，缺省未定义)
                    If lng总住院天数 = 0 Then lng总住院天数 = Get总数(strPatiPage, lng总人次, 2, strTablePage, intMaxPage, varPage)
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总住院天数, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总住院天数, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 27 '抗菌药物费用比
                    If blnGeted总费用 = False Then lng总费用 = Get总数(strPatiPage, lng总人次, 25, strTablePage, intMaxPage, varPage): blnGeted总费用 = True
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总费用, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总费用, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 28 '耗材费用比
                    If blnGeted总费用 = False Then lng总费用 = Get总数(strPatiPage, lng总人次, 25, strTablePage, intMaxPage, varPage): blnGeted总费用 = True
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总费用, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总费用, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                
                Case 29 '检查费用比
                    If blnGeted总费用 = False Then lng总费用 = Get总数(strPatiPage, lng总人次, 25, strTablePage, intMaxPage, varPage): blnGeted总费用 = True
                    If Len(strPatiPage) <= 4000 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总费用, mlng路径ID)
                    Else
                        Call FuncReDoSQLNum(strSQL, 2, 3)
                        strTempSQL = strTablePage
                        Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 2) '整体向后移动2位
                        strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([3]))"), strTempSQL)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总费用, mlng路径ID, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                            CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
                    End If
                    If gcnOracle.Errors.count = 0 Then
                        If rsTmp.RecordCount > 0 Then .TextMatrix(i, col结果) = "" & rsTmp.Fields(0).Value
                    End If
                
                End Select
            
                If gcnOracle.Errors.count <> 0 Then
                    .Cell(flexcpData, i, col备注) = CStr(gcnOracle.Errors(0).Description)
                    .Cell(flexcpBackColor, i, col备注) = &HC0FFFF
                    gcnOracle.Errors.Clear
                    stbThis.Panels(2).Text = "第" & i & "行读取数据出错：" & CStr(gcnOracle.Errors(0).Description)
                    lngErrCnt = lngErrCnt + 1
                    Me.Refresh
                End If
                
                '没有定义SQL的单元格取数
            Else
                Select Case .RowData(i)
                    Case 18 '住院患者总人数
                        .TextMatrix(i, col结果) = lng总人数
                    Case 26 '单病种日均费用
                        If blnGeted总费用 = False Then lng总费用 = Get总数(strPatiPage, lng总人次, 25, strTablePage, intMaxPage, varPage): blnGeted总费用 = True
                        If lng总住院天数 = 0 Then lng总住院天数 = Get总数(strPatiPage, lng总人次, 2, strTablePage, intMaxPage, varPage)
                        .TextMatrix(i, col结果) = Format(lng总费用 / lng总住院天数, "#######0.00")
                End Select
            End If
        Next
    End With
    
    If lngErrCnt <> 0 Then
        stbThis.Panels(2).Text = "共" & lngErrCnt & "行数据读取出错,鼠标移到黄色单元格可查看详细错误。"
    Else
        stbThis.Panels(2).Text = "已完成数据读取操作。"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function Get总数(ByVal strPatiPage As String, ByVal lng总人次 As Long, ByVal lngRow As Long, ByVal strTablePage As String, _
    ByVal intMaxPage As Integer, ByVal varPage As Variant) As Long
'功能：获取总费用(第25行)，总住院天数(第2行)
'参数：strPatiPage:病人及主页ID列表:1121:1,1122:1
'      lngRow:报表行号，需转换为表格行号
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strTempSQL As String

    On Error Resume Next
    With vsItem
        For i = .FixedRows To .Rows - 1
            If .RowData(i) = lngRow Then Exit For
        Next
    
        strSQL = .Cell(flexcpData, i, col结果)
        If strSQL <> "" Then
            If Len(strPatiPage) <= 4000 Then
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatiPage, lng总人次)
            Else
                '参数号调整：将未超长的参数排在最前面
                Call FuncReDoSQLNum(strSQL, 2, 2)
                strTempSQL = strTablePage
                Call FuncMoveSQLNum(strTempSQL, 1, intMaxPage, 1) '整体向后移动一位
                strSQL = Replace(UCase(strSQL), UCase("Table(f_Num2list2([2]))"), strTempSQL)
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng总人次, CStr(varPage(0)), CStr(varPage(1)), CStr(varPage(2)), CStr(varPage(3)), _
                    CStr(varPage(4)), CStr(varPage(5)), CStr(varPage(6)), CStr(varPage(7)), CStr(varPage(8)), CStr(varPage(9)))
            End If
            If gcnOracle.Errors.count = 0 Then
                If rsTmp.RecordCount > 0 Then
                    If lngRow = 25 Or lngRow = 2 Then
                        If rsTmp.Fields.count > 1 Then Get总数 = Val("" & rsTmp.Fields(1).Value)  '获取总费用，总住院天数,固定取第2个字段
                    Else
                        Get总数 = Val("" & rsTmp.Fields(0).Value)
                    End If
                End If
            Else
                .Cell(flexcpData, i, col备注) = CStr(gcnOracle.Errors(0).Description)
                .Cell(flexcpBackColor, i, col备注) = &HC0FFFF
                gcnOracle.Errors.Clear
            End If
        End If
    End With
End Function

Private Sub vsItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsItem
        If .MouseCol >= .FixedCols And .MouseRow >= .FixedRows Then
            Dim strErr As String
            If .Col = col备注 Then
                strErr = .Cell(flexcpData, .MouseRow, col备注)
                If strErr <> "" Then
                    Call zlCommFun.ShowTipInfo(.Hwnd, strErr)
                Else
                    Call zlCommFun.ShowTipInfo(.Hwnd, "")
                End If
            Else
                Call zlCommFun.ShowTipInfo(.Hwnd, "")
            End If
        End If
    End With
End Sub

Private Function FuncGetTable(ByVal strPara As String, ByVal bytFunc As Byte, ByRef strTableOut As String, ByRef intMaxIdx As Integer) As Variant
'功能：对于动态内存表的绑定参数超长的处理
'参数：strPar 参数串 大于4000,拆分分隔符默认是","
'   bytFunc=0 内存表f_Num2list;bytFunc=1 :f_Num2list2
'返回：一个字符串数组，最多10个元素
'    strTableOut=返回与动态内存表等效的SQL语句(UNION ALL 相连)
'   intMaxIdx = 返回拆分之后得到最大参数序号
    Dim varPara As Variant
    Dim strParTable As String
    
    varPara = Array()
    
    If bytFunc = 0 Then
        strParTable = "Select Column_Value From Table(f_Num2list([1]))"
    Else
        strParTable = "Select C1, C2 From Table(f_Num2list2([1]))"
    End If
    varPara = GetParTable(strPara, strParTable, strTableOut, intMaxIdx)
    strTableOut = "(" & strTableOut & ")"
                     
    FuncGetTable = varPara
End Function


Private Sub FuncMoveSQLNum(ByRef strSQL As String, ByVal intBegin As Integer, ByVal intEnd As Integer, ByVal intMoveLen As Integer)
'功能:将SQL中参数序号整体向后移动或向前移动
'参数:intBegin,intEnd=调整的参数序号的闭区间[intBegin,intEnd]
'intMoveLen=偏移量 >0向后移动,<0向前移动
    Dim i As Integer
    If intMoveLen > 0 Then
    i = intEnd
        Do While i >= intBegin
            strSQL = Replace(strSQL, "[" & i & "]", "[" & (i + intMoveLen) & "]")
            i = i - 1
        Loop
    Else
        For i = intBegin To intEnd
             strSQL = Replace(strSQL, "[" & i & "]", "[" & (i + intMoveLen) & "]")
        Next
    End If
End Sub

Private Sub FuncReDoSQLNum(ByRef strSQL As String, ByVal intBegin As Integer, ByVal intEnd As Integer)
'功能:将SQL中参数值长度超过4000的参数序号调整到该SQL中最大的参数序号
'参数:
'intBegin:需要调整参数起始位置
'intEnd:需要调整参数末尾位置（intEnd>=intBegin）
'intBegin-1:该参数值对应长度超过4000 ;（intEnd + 1）用于临时存储
'返回:调整后的SQL
    strSQL = Replace(strSQL, "[" & (intBegin - 1) & "]", "[" & (intEnd + 1) & "]")
    Call FuncMoveSQLNum(strSQL, intBegin, (intEnd + 1), -1)
End Sub
