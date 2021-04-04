VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmNurseFileMan 
   Caption         =   "护理文件管理"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10335
   Icon            =   "frmNurseFileMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10335
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   5025
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   660
      Width           =   8400
      _Version        =   589884
      _ExtentX        =   14817
      _ExtentY        =   8864
      _StockProps     =   0
      BorderStyle     =   1
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1590
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNurseFileMan.frx":5162
            Key             =   "体温单"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNurseFileMan.frx":5874
            Key             =   "记录单"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic归档 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1080
      Picture         =   "frmNurseFileMan.frx":5F86
      ScaleHeight     =   345
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   90
      Width           =   375
   End
   Begin VB.PictureBox pic病人 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2340
      ScaleHeight     =   225
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   150
      Width           =   1365
      Begin VB.ComboBox cbo病人 
         BackColor       =   &H00EAFFFF&
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   -30
         Width           =   1425
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmNurseFileMan.frx":6688
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15319
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   540
      Left            =   8460
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   1095
      _cx             =   1931
      _cy             =   952
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
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   510
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmNurseFileMan.frx":6F1A
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   60
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmNurseFileMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrSQL As String
Private mblnSaved As Boolean            '进入本模块后是否保存过数据
Private mlng病人id As Long
Private mlng主页id As Long
Private mint婴儿 As Integer
Private mlng科室ID As Long
Private mstr科室 As String
Private mblnPigeonhole As Boolean       '归档
Private mblnFileEnd As Boolean          '文件结束
Private mblnPrintMerge As Boolean       '合并打印
Private mintNORule As Integer           '护理文件页码规则:住院期间统一编号时不允许设置文件为合并打印
Private Enum COLDef
    c_图标
    c_文件名称
    c_文件来源
    c_开始时间
    c_结束时间
    c_续打记录单
    c_创建人
    c_创建时间
End Enum

'绑定快捷键时,ID值如大于无符号整型的取值范围则无法绑定,也就是0-65535
Private Const conMenu_Add As Long = 32761
Private Const conMenu_Modify As Long = 32762
Private Const conMenu_Delete As Long = 32763
Private Const conMenu_FileEnd As Long = 32764
Private Const conMenu_FileRestore As Long = 32765
Private Const conMenu_PrintMerge As Long = 32766
Private Const conMenu_PrintSingle As Long = 32767


Public Function ShowEditor(ByVal lng病人id As Long, ByVal lng主页id As Long, ByVal intBaby As Integer) As Boolean
    mlng病人id = lng病人id
    mlng主页id = lng主页id
    mint婴儿 = intBaby
    mblnSaved = False
    Me.Show 1
    ShowEditor = mblnSaved
End Function

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objTool As CommandBar
    Dim objMenu As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    
    'cbsMain
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
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = imgPublic.Icons
    
    '菜单项
    cbsMain.ActiveMenuBar.Title = "菜单栏"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Add, "新增(&A)"): objControl.IconId = 1
        Set objControl = .Add(xtpControlButton, conMenu_Modify, "修改(&M)"): objControl.IconId = 2
        Set objControl = .Add(xtpControlButton, conMenu_Delete, "删除(&D)"): objControl.IconId = 3
        Set objControl = .Add(xtpControlButton, conMenu_FileEnd, "标记结束(&E)"): objControl.IconId = 4: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_FileRestore, "继续记录(&C)"): objControl.IconId = 5
        Set objControl = .Add(xtpControlButton, conMenu_PrintMerge, "合并打印(&G)"): objControl.IconId = 6: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_PrintSingle, "独立打印(&L)"): objControl.IconId = 7
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        objControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        objControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        objControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        objControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        objControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): objControl.BeginGroup = True
    End With
    '增加归档标志
    Set objCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Option, "归档")
    objCustom.Handle = Me.pic归档.hWnd
    objCustom.Flags = xtpFlagRightAlign
    cbsMain(1).EnableDocking xtpFlagHideWrap + xtpFlagStretched

    '工具栏定义
    '-----------------------------------------------------
    Set objTool = cbsMain.Add("工具栏", xtpBarTop)      '固有
    objTool.EnableDocking xtpFlagStretched
    
    objTool.Closeable = False
    With objTool.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Add, "增加"): objControl.IconId = 1: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Modify, "修改"): objControl.IconId = 2: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Delete, "删除"): objControl.IconId = 3: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_FileEnd, "结束"): objControl.IconId = 4: objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True: objControl.ToolTipText = "标记当前文件结束"
        Set objControl = .Add(xtpControlButton, conMenu_FileRestore, "取消"): objControl.IconId = 5: objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "取消当前文件的结束标志"
        Set objControl = .Add(xtpControlButton, conMenu_PrintMerge, "合并"): objControl.IconId = 6: objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True: objControl.ToolTipText = "指定格式相同的两份文件为合并打印"
        Set objControl = .Add(xtpControlButton, conMenu_PrintSingle, "独立"): objControl.IconId = 7: objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "当前文件设定为独立打印"
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.Style = xtpButtonIconAndCaption
    End With
    '特殊处理
    '-----------------------------------------------------
    '工具栏右侧病区下拉框选择
    With objTool.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_View_Find, "病人")
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "病人")
        objCustom.Handle = Me.pic病人.hWnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Add               '增加护理文件
        .Add 0, vbKeyDelete, conMenu_Delete              '删除护理文件
        .Add 0, vbKeyF1, conMenu_Help_Help               '帮助
        .Add 0, vbKeyF5, conMenu_View_Refresh
    End With
End Sub

Private Sub cbo病人_Click()
    On Error GoTo errHand
    Dim objItem As ReportRecordItem
    Dim objRecord As ReportRecord
    Dim objRpt As ReportControl
    Dim rsTemp As New ADODB.Recordset

    '显示指定病人的护理文件列表
    mstrSQL = " Select A.ID,A.文件名称, B.名称 AS 文件来源,A.开始时间,A.结束时间,A.创建人,A.创建时间,A.归档人,C.文件名称 AS 续打文件,B.保留 " & _
              " From 病人护理文件 A,病历文件列表 B,病人护理文件 C" & _
              " Where A.格式ID=B.ID And A.病人ID=[1] And A.主页ID=[2] And A.婴儿=[3] And A.续打ID=C.ID(+)" & _
              " Order by B.保留,A.开始时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "显示指定病人的护理文件列表", mlng病人id, mlng主页id, Val(cbo病人.ItemData(cbo病人.ListIndex)))
    
    mblnPigeonhole = False
    rptList.Records.DeleteAll
    With rsTemp
        If .RecordCount <> 0 Then
            mblnPigeonhole = (NVL(!归档人) <> "")
        End If

        '将记录加入报表控件
        Do While Not .EOF
            Set objRecord = Me.rptList.Records.Add()
            objRecord.Tag = CStr(!ID)
            Set objItem = objRecord.AddItem(""): objItem.Icon = Val(!保留) + 1
            Set objItem = objRecord.AddItem(CStr(!文件名称))
            objItem.Caption = CStr(!文件名称)
            Set objItem = objRecord.AddItem(CStr(!文件来源))
            objItem.Caption = CStr(!文件来源)
            Set objItem = objRecord.AddItem(CStr(Format(!开始时间, "yyyy-MM-dd HH:mm:ss")))
            objItem.Caption = CStr(Format(!开始时间, "yyyy-MM-dd HH:mm:ss"))
            Set objItem = objRecord.AddItem(CStr(Format(!结束时间, "yyyy-MM-dd HH:mm:ss")))
            objItem.Caption = CStr(Format(!结束时间, "yyyy-MM-dd HH:mm:ss"))
            Set objItem = objRecord.AddItem(CStr(NVL(!续打文件)))
            objItem.Caption = CStr(NVL(!续打文件))
            Set objItem = objRecord.AddItem(CStr(!创建人))
            objItem.Caption = CStr(!创建人)
            Set objItem = objRecord.AddItem(CStr(Format(!创建时间, "yyyy-MM-dd HH:mm:ss")))
            objItem.Caption = CStr(Format(!创建时间, "yyyy-MM-dd HH:mm:ss"))
            
            .MoveNext
        Loop
    End With
    rptList.Populate

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(vsfPrint, rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsfPrint
    
    objPrint.Title.Text = "护理文件清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub LoadPati()
    Dim strName As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '读取病人当前病区
    mstrSQL = " Select B.ID,B.名称" & _
              " From 病案主页 A,部门表 B" & _
              " Where A.病人ID=[1] And A.主页ID=[2] And A.出院科室ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "读取病人当前病区", mlng病人id, mlng主页id)
    mlng科室ID = rsTemp!ID
    mstr科室 = rsTemp!名称
    
    '读取病人身份
    mstrSQL = "" & _
            "SELECT 病人ID,0 AS 序号,姓名,性别 FROM 病人信息 WHERE 病人ID=[1]" & vbNewLine & _
            "UNION" & vbNewLine & _
            "SELECT 病人ID,序号,婴儿姓名 AS 姓名,婴儿性别 AS 性别 FROM 病人新生儿记录 WHERE 病人ID=[1] AND 主页ID=[2]" & vbNewLine & _
            "ORDER BY 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "读取病人身份", mlng病人id, mlng主页id)
    
    With rsTemp
        cbo病人.Clear
        Do While Not .EOF
            If !序号 = 0 Then strName = !姓名
            cbo病人.AddItem IIf(IsNull(!姓名), strName & "之子" & !序号, !姓名)
            cbo病人.ItemData(cbo病人.NewIndex) = !序号
            If mint婴儿 = !序号 Then cbo病人.ListIndex = .AbsolutePosition - 1
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitRpt()
    Dim objCol As ReportColumn
    With rptList
        .Columns.DeleteAll
        Set objCol = .Columns.Add(c_图标, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_文件名称, "记录单名称", 120, True)
        Set objCol = .Columns.Add(c_文件来源, "文件来源", 120, True)
        Set objCol = .Columns.Add(c_开始时间, "开始时间", 130, True)
        Set objCol = .Columns.Add(c_结束时间, "结束时间", 130, True)
        Set objCol = .Columns.Add(c_续打记录单, "续打记录单", 120, True)
        Set objCol = .Columns.Add(c_创建人, "创建人", 70, True)
        Set objCol = .Columns.Add(c_创建时间, "创建时间", 130, True)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Sortable = True
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            '.HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有文件..."
        End With
        .TabStop = True
        .PreviewMode = False
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgList
    End With
End Sub

'控件事件##############################################################################################################

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim cbrControl As Object

    On Error GoTo errHand

    Select Case Control.ID
        Case conMenu_File_PrintSet

            Call zlPrintSet

        Case conMenu_File_Preview

            Call zlRptPrint(2)

        Case conMenu_File_Print

            Call zlRptPrint(1)

        Case conMenu_File_Excel

            Call zlRptPrint(3)

        Case conMenu_Add
            If frmNurseFileEdit.ShowEditor(mlng病人id, mlng主页id, Me.cbo病人.ItemData(Me.cbo病人.ListIndex), mlng科室ID, mstr科室) Then
                mblnSaved = True
                Call cbo病人_Click
            End If
        Case conMenu_Modify
            If frmNurseFileEdit.ShowEditor(mlng病人id, mlng主页id, Me.cbo病人.ItemData(Me.cbo病人.ListIndex), mlng科室ID, mstr科室, Val(rptList.FocusedRow.Record.Tag)) Then
                mblnSaved = True
                Call cbo病人_Click
            End If
        Case conMenu_Delete
            If MsgBox("你确定要删除" & rptList.FocusedRow.Record.Item(c_文件名称).Caption & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If MsgBox("该文件所有的护理数据也将一并删除，请再次确认是否删除！", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Call zlDatabase.ExecuteProcedure("ZL_病人护理文件_DELETE(" & Val(rptList.FocusedRow.Record.Tag) & ")", "删除文件")
            mblnSaved = True
            Call cbo病人_Click
        Case conMenu_FileEnd
            gstrSQL = "ZL_病人护理文件_STATE(" & Val(rptList.FocusedRow.Record.Tag) & ",1,sysdate)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "标记文件结束")
            Call cbo病人_Click
        Case conMenu_FileRestore
            gstrSQL = "ZL_病人护理文件_STATE(" & Val(rptList.FocusedRow.Record.Tag) & ",1,NULL)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "取消文件结束")
            Call cbo病人_Click
        Case conMenu_PrintMerge
            If frmNurseFileMerge.ShowEditor(Val(rptList.FocusedRow.Record.Tag)) Then cbo病人_Click
        Case conMenu_PrintSingle
            gstrSQL = "ZL_病人护理文件_STATE(" & Val(rptList.FocusedRow.Record.Tag) & ",2,NULL)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "取消合并打印")
            Call cbo病人_Click
        Case conMenu_View_ToolBar_Button

            cbsMain(2).Visible = Not cbsMain(2).Visible
            cbsMain.RecalcLayout

        Case conMenu_View_ToolBar_Text

            For Each cbrControl In cbsMain(2).Controls
                If cbrControl.Type = xtpControlButton Then
                    cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next

            cbsMain.RecalcLayout

        Case conMenu_View_StatusBar

            stbThis.Visible = Not stbThis.Visible
            cbsMain.RecalcLayout

        Case conMenu_View_Refresh
            Call cbo病人_Click

        Case conMenu_Help_Help

            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))

        Case conMenu_Help_About

            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)

        Case conMenu_Help_Web_Home

            Call zlHomePage(Me.hWnd)

        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hWnd)

        Case conMenu_Help_Web_Mail

            Call zlMailTo(Me.hWnd)

        Case conMenu_File_Exit
            Unload Me
            Exit Sub
            Exit Sub
    End Select

    cbsMain.RecalcLayout

    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (rptList.Records.Count > 0)
    Case conMenu_Add
        Control.Enabled = Not mblnPigeonhole
    Case conMenu_Modify, conMenu_Delete
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                '只能修改、删除自己创建的文件
                Control.Enabled = (rptList.FocusedRow.Record.Item(c_创建人).Value = gstrUserName) And Not mblnPigeonhole
            End If
        End If
    Case conMenu_FileEnd
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnFileEnd And Not mblnPigeonhole And (rptList.FocusedRow.Record.Item(c_图标).Icon > 0)
            End If
        End If
    Case conMenu_FileRestore
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = mblnFileEnd And Not mblnPigeonhole And (rptList.FocusedRow.Record.Item(c_图标).Icon > 0)
            End If
        End If
    Case conMenu_PrintMerge
        Control.Enabled = False
        If mintNORule = 1 Then Exit Sub         '护理文件页码规则为住院期间统一编号时,不允许使用合并功能
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnPrintMerge And (rptList.FocusedRow.Record.Item(c_图标).Icon > 0)
            End If
        End If
    Case conMenu_PrintSingle
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = mblnPrintMerge And (rptList.FocusedRow.Record.Item(c_图标).Icon > 0)
            End If
        End If
    Case conMenu_View_Option    '归档标志
        Control.Visible = mblnPigeonhole
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsMain(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub Form_Load()
    mintNORule = zlDatabase.GetPara("护理文件页码规则", glngSys, 1255, 0)
    
    Call MainDefCommandBar
    Call InitRpt
    Call LoadPati
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    On Error Resume Next
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With rptList
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight
        .Height = lngBottom - lngTop - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If rptList.Records.Count = 0 Then Exit Sub
    If rptList.FocusedRow Is Nothing Then Exit Sub
    If mblnPigeonhole Then Exit Sub
    
    Call cbsMain_Execute(cbsMain.FindControl(, conMenu_Modify))
End Sub

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptList_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub rptList_SelectionChanged()
    If rptList.Records.Count = 0 Then Exit Sub
    If rptList.FocusedRow Is Nothing Then Exit Sub
    
    mblnFileEnd = (rptList.FocusedRow.Record.Item(c_结束时间).Caption <> "")
    mblnPrintMerge = (rptList.FocusedRow.Record.Item(c_续打记录单).Caption <> "")
End Sub
