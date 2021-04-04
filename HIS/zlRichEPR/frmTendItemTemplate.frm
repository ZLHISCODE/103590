VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmTendItemTemplate 
   Caption         =   "护理项目模板"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8685
   Icon            =   "frmTendItemTemplate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   8685
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList imgrpt 
      Left            =   4050
      Top             =   2580
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
            Picture         =   "frmTendItemTemplate.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendItemTemplate.frx":D0B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   5820
      ScaleHeight     =   3195
      ScaleWidth      =   2505
      TabIndex        =   2
      Top             =   930
      Width           =   2505
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   3105
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2400
         _cx             =   4233
         _cy             =   5477
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
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
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   480
      ScaleHeight     =   3945
      ScaleWidth      =   4845
      TabIndex        =   1
      Top             =   750
      Width           =   4845
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2040
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1995
         _Version        =   589884
         _ExtentX        =   3519
         _ExtentY        =   3598
         _StockProps     =   0
         BorderStyle     =   2
         ShowGroupBox    =   -1  'True
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendItemTemplate.frx":13916
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10239
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      DesignerControls=   "frmTendItemTemplate.frx":141A8
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   690
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTendItemTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'窗口级变量申明########################################################################################################
Private Enum mHeadCol
    图标
    模板名称
    适用护理等级
    护理等级
    科室
    科室ID
End Enum
Private Enum mDetailCol
    序号
    项目名称
End Enum

Private mstrSel As String           '记录当前选择项目的信息,便于新增,修改后定位,找不到则定位到后继项目上;如果为空,表示定位的第一个项目上
Private mlng科室ID As Long          '当前操作员所属的缺省科室ID
Private mstr科室ID As String        '当前操作员所属科室ID
Private mstrPrivs As String         '当前使用者权限串
Private mblnStartUp As Boolean
Private mstrSQL As String
Private mstrDeptID As String        '当前操作员所属科室ID

'自定义过程/函数申明###################################################################################################

Public Sub ShowME(ByVal objParent As Object, ByVal strPrivs As String)
    mstrPrivs = strPrivs
    Me.Show 1, objParent
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If rptList.FocusedRow Is Nothing Then Exit Sub
    If rptList.FocusedRow.Record Is Nothing Then Exit Sub

    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow

    Set objPrint.Body = vsfDetail

    objPrint.Title.Text = rptList.FocusedRow.Record.Item(模板名称).Value
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("科室:" & rptList.FocusedRow.Record.Item(科室).Value)
    Call objPrint.UnderAppRows.Add(objAppRow)
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("适用护理等级:" & rptList.FocusedRow.Record.Item(适用护理等级).Value)
    Call objPrint.UnderAppRows.Add(objAppRow)
    
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

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新状态栏显示信息
    '------------------------------------------------------------------------------------------------------------------
    stbThis.Panels(2).Text = "共有 " & vsfDetail.Rows - 1 & " 个护理记录项目！"
End Sub

Private Function zlMenuClick(ByVal strMenuItem As String, Optional ByVal strParam As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：功能处理
    '------------------------------------------------------------------------------------------------------------------
    Dim arrData
    Dim blnSel As Boolean
    Dim str护理等级 As String
    Dim intRow As Integer, intRows As Integer
    Dim rptRecord As ReportRecord, rptRecordItem As ReportRecordItem
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    Select Case strMenuItem
    Case "初始化"
        '提取当前操作员所属的科室ID
        gstrSQL = " Select B.ID,B.编码,B.名称 " & _
                  " From 部门性质说明 A,部门表 B,部门人员 C" & _
                  " Where A.工作性质='临床' And A.服务对象 IN (2,3) And A.部门ID=B.ID" & _
                  " And B.ID=C.部门ID And C.人员ID=[1]" & _
                  " UNION " & _
                  " Select B.ID,B.编码,B.名称 " & _
                  " From 部门性质说明 A,部门表 B,病区科室对应 C" & _
                  " Where A.工作性质='临床' And A.服务对象 IN (2,3) And A.部门ID=B.ID And B.ID=C.科室ID And C.病区ID=[2]"
        gstrSQL = " Select Distinct ID,编码,名称 From (" & gstrSQL & ") Order by 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId, glngDeptId)
        With rsTemp
            mstr科室ID = ""
            Do While Not .EOF
                mstr科室ID = mstr科室ID & "," & !ID
                If mlng科室ID = 0 Then mlng科室ID = !ID
                .MoveNext
            Loop
        End With
        
    Case "读取数据"
        Call InitGird
        rptList.Records.DeleteAll
        
        '提取所有模板
        mstrSQL = " Select Distinct B.ID,NVL(B.名称,'') AS 科室,模板名称,护理等级 " & _
                  " From 护理项目模板 A,部门表 B" & _
                  " Where A.科室ID=B.ID(+)" & _
                  " Order by NVL(B.名称,''),Decode(护理等级,-1,5,护理等级) "
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
        If rsTemp.BOF = False Then
            Do While Not rsTemp.EOF
                Set rptRecord = rptList.Records.Add()
                Set rptRecordItem = rptRecord.AddItem("")
                rptRecordItem.Icon = IIf(Val(NVL(rsTemp!ID)) > 0, 1, 0)
                rptRecord.AddItem rsTemp.Fields("模板名称").Value
                
                Select Case rsTemp!护理等级
                Case -1
                    str护理等级 = "批量录入模板"
                Case 0
                    str护理等级 = "特级护理录入模板"
                Case 1
                    str护理等级 = "一级护理录入模板"
                Case 2
                    str护理等级 = "二级护理录入模板"
                Case 3
                    str护理等级 = "三级护理录入模板"
                End Select
                rptRecord.AddItem str护理等级
                rptRecord.AddItem rsTemp.Fields("护理等级").Value
                rptRecord.AddItem rsTemp.Fields("科室").Value
                rptRecord.AddItem NVL(rsTemp.Fields("ID").Value, 0)
                
                rsTemp.MoveNext
            Loop
        End If
        rptList.Populate
        
        '定位项目
        On Error Resume Next
        If mstrSel = "" Then
            If rptList.Rows.Count > 0 Then Set rptList.FocusedRow = rptList.Rows(1)
        Else
            arrData = Split(mstrSel, "|")
            intRows = rptList.Rows.Count
            For intRow = 1 To intRows
                If Not rptList.Rows(intRow - 1).Record Is Nothing Then
                    If Val(rptList.Rows(intRow - 1).Record.Item(护理等级).Value) = arrData(0) And Val(rptList.Rows(intRow - 1).Record.Item(科室ID).Value) = arrData(1) Then
                        blnSel = True
                        Set rptList.FocusedRow = rptList.Rows(intRow - 1)
                        Exit For
                    End If
                End If
            Next
            If blnSel = False Then
                '没找到,可能删除掉了,直接定位到其后一个记录
                If Val(arrData(2)) <= rptList.Rows.Count Then
                    On Error Resume Next
                    rptList.Rows(arrData(2)).Selected = True
                Else
                    '说明删除的是最后一条,定位在最后一条上
                    If rptList.Rows.Count > 0 Then Set rptList.FocusedRow = rptList.Rows(rptList.Rows.Count).Selected
                End If
            End If
        End If
        
    Case "读取模板内容"
        Call InitGird
        
        If rptList.FocusedRow Is Nothing Then Exit Function
        If rptList.FocusedRow.Record Is Nothing Then Exit Function
        gstrSQL = " Select B.项目序号,B.项目名称 From 护理项目模板 A,护理记录项目 B " & _
                  " Where A.项目序号=B.项目序号 And A.科室ID =[1] And A.护理等级=[2]" & _
                  " Order by A.排列序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(rptList.FocusedRow.Record.Item(科室ID).Value), CInt(rptList.FocusedRow.Record.Item(护理等级).Value))
        
        With rsTemp
            Do While Not .EOF
                If .AbsolutePosition > vsfDetail.Rows - 1 Then vsfDetail.Rows = vsfDetail.Rows + 1
                vsfDetail.TextMatrix(.AbsolutePosition, 序号) = .Fields("项目序号").Value
                vsfDetail.TextMatrix(.AbsolutePosition, 项目名称) = .Fields("项目名称").Value
                rsTemp.MoveNext
            Loop
        End With
        Call RefreshStateInfo
    End Select
    '------------------------------------------------------------------------------------------------------------------

    cbsThis.RecalcLayout
    Call RefreshStateInfo

    zlMenuClick = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitRpt(Optional ByVal intState As Integer = 0)
    '0表示两个表格都刷新;1-只刷新明细表
    Dim rptCol As ReportColumn
    
    With rptList
        Set rptCol = .Columns.Add(图标, "", 20, False)
        rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        
        Set rptCol = .Columns.Add(模板名称, "模板名称", 250, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(适用护理等级, "适用护理等级", 111, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(护理等级, "护理等级", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(科室, "科室", 60, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(科室ID, "科室ID", 0, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        
        .SetImageList imgrpt
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GridLineColor = RGB(225, 225, 225)
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
        End With
        .PreviewMode = True
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(科室)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(护理等级)
    End With
    
End Sub

Private Sub InitGird()
    With vsfDetail
        .Clear
        .Rows = 2: .Cols = 2
        .TextMatrix(0, 序号) = "序号"
        .TextMatrix(0, 项目名称) = "项目名称"
        .ColWidth(序号) = 800
        .ColWidth(项目名称) = 2000
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub

Private Function InitMenuBar() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化菜单、工具栏
    '------------------------------------------------------------------------------------------------------------------
    Dim cbrMenuBar As Object
    Dim obj As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrToolBar As CommandBar
    Dim objExtendedBar As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): cbrControl.BeginGroup = True
    End With
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
               
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '读取发布到该模块的报表:因为是一次性读取,全局变量可用
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
End Function

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

'控件事件##############################################################################################################

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
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

        Case conMenu_View_ToolBar_Button

            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout

        Case conMenu_View_ToolBar_Text

            For Each cbrControl In cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next

            cbsThis.RecalcLayout

        Case conMenu_View_StatusBar

            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout

        Case conMenu_Edit_NewItem
            '新增项目
            frmTendItemTemplateEdit.mstrPrivs = mstrPrivs
            strKey = frmTendItemTemplateEdit.ShowEditor(Me, mlng科室ID, "", 9)
            If strKey = "" Then Exit Sub
            mstrSel = strKey
            Call zlMenuClick("读取数据")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify

            If rptList.FocusedRow Is Nothing Then Exit Sub
            If rptList.FocusedRow.Record Is Nothing Then Exit Sub

            '修改项目
            frmTendItemTemplateEdit.mstrPrivs = mstrPrivs
            If frmTendItemTemplateEdit.ShowEditor(Me, rptList.FocusedRow.Record.Item(科室ID).Value, rptList.FocusedRow.Record.Item(模板名称).Value, CInt(rptList.FocusedRow.Record.Item(护理等级).Value)) <> "" Then Call zlMenuClick("读取数据")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            '删除项目
            If rptList.FocusedRow Is Nothing Then Exit Sub
            If rptList.FocusedRow.Record Is Nothing Then Exit Sub

            If MsgBox("你真的要删除“" & rptList.FocusedRow.Record.Item(模板名称).Value & "”？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Call zlDatabase.ExecuteProcedure("zl_护理项目模板_Delete(" & rptList.FocusedRow.Record.Item(科室ID).Value & "," & CInt(rptList.FocusedRow.Record.Item(护理等级).Value) & ")", "删除模板")
            Call zlMenuClick("读取数据")
        '--------------------------------------------------------------------------------------------------------------

        Case conMenu_View_Refresh
            Call zlMenuClick("读取数据")

        Case conMenu_Help_Help

            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))

        Case conMenu_Help_About

            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)

        Case conMenu_Help_Web_Home

            Call zlHomePage(Me.hwnd)

        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hwnd)

        Case conMenu_Help_Web_Mail

            Call zlMailTo(Me.hwnd)

        Case conMenu_File_Exit
            Unload Me
            Exit Sub
        Case Else
            '执行发布到当前模块的报表
'            Dim lng项目序号 As Long, str项目名称 As String
'            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
'                If rptList.SelectedRows.Count > 0 Then
'                    If Not rptList.SelectedRows(0).GroupRow Then
'                        lng项目序号 = Val(rptList.SelectedRows(0).Record(mCol.项目序号).Value)
'                        str项目名称 = rptList.SelectedRows(0).Record(mCol.项目名称).Value
'                    End If
'                End If
'                If str项目名称 <> "" Then
'                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "项目序号=" & lng项目序号, "项目名称=" & str项目名称)
'                Else
'                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
'                End If
'            End If
            Exit Sub
    End Select

    cbsThis.RecalcLayout
    Call RefreshStateInfo

    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)

    If stbThis.Visible Then Bottom = stbThis.Height
    
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (vsfDetail.Rows - 1 > 0)
    Case conMenu_Edit_NewItem
        Control.Enabled = (InStr(1, mstrPrivs, "护理模板") > 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                '首先要有护理模板的编辑权限
                Control.Enabled = (InStr(1, mstrPrivs, "护理模板") > 0)
                If Control.Enabled And Val(rptList.FocusedRow.Record.Item(科室ID).Value) <> 0 Then
                    '操作员如果有编辑其它科室模板的权限,则允许修改或删除,否则不允许
                    Control.Enabled = (InStr(1, ";" & mstrPrivs & ";", ";编辑其它科室模板;") > 0) Or (InStr(1, mstr科室ID & ",", "," & rptList.FocusedRow.Record.Item(科室ID).Value & ",") <> 0)
                End If
            End If
        End If
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size
        Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picHead.hwnd
    Case 2
        Item.Handle = picDetail.hwnd
    End Select
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    
    mstrSel = ""
    mblnStartUp = True
    
    Call InitCommonControls
    Call InitMenuBar
    Call InitRpt
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.UseSplitterTracker = False '实时拖动
    dkpMan.Options.AlphaDockingContext = True
    dkpMan.Options.CloseGroupOnButtonClick = True
    dkpMan.Options.HideClient = True
    dkpMan.SetCommandBars cbsThis
    Set objPane = dkpMan.CreatePane(1, 5400, 0, DockLeftOf, Nothing): objPane.Title = "主单": objPane.Options = PaneNoCaption
    Set objPane = dkpMan.CreatePane(2, vsfDetail.Width, vsfDetail.Height, DockRightOf, Nothing): objPane.Title = "子单": objPane.Options = PaneNoCaption
    
    Call RestoreWinState(Me, App.ProductName)
    
    mblnStartUp = False
    Call zlMenuClick("初始化")
    Call zlMenuClick("读取数据")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picDetail_Resize()
    With picDetail
        vsfDetail.Left = 0
        vsfDetail.Top = 0
        vsfDetail.Width = picDetail.Width
        vsfDetail.Height = picDetail.Height
    End With
End Sub

Private Sub picHead_Resize()
    With picHead
        rptList.Left = 0
        rptList.Top = 0
        rptList.Width = picHead.Width
        rptList.Height = picHead.Height
    End With
End Sub

Private Sub rptList_SelectionChanged()
    If rptList.FocusedRow Is Nothing Then Exit Sub
    If rptList.FocusedRow.Record Is Nothing Then Exit Sub
    
    mstrSel = Val(rptList.FocusedRow.Record.Item(护理等级).Value) & "|" & Val(rptList.FocusedRow.Record.Item(科室ID).Value) & "|" & rptList.FocusedRow.Index
    Call zlMenuClick("读取模板内容")
End Sub
