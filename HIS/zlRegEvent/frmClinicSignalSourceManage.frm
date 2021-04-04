VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmClinicSignalSourceManage 
   BorderStyle     =   0  'None
   Caption         =   "临床号源管理"
   ClientHeight    =   9975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   9360
      MaxLength       =   100
      TabIndex        =   5
      Top             =   390
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.PictureBox picDetailList 
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   3930
      ScaleHeight     =   3075
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   3390
      Width           =   3675
      Begin zl9RegEvent.ClinicPlanDetailPages CPDPages 
         Height          =   2535
         Left            =   660
         TabIndex        =   4
         Top             =   180
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   4471
         BackColor       =   -2147483628
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
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Left            =   2790
      ScaleHeight     =   2445
      ScaleWidth      =   6900
      TabIndex        =   0
      Top             =   870
      Width           =   6900
      Begin XtremeReportControl.ReportControl rptData 
         Height          =   1425
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3705
         _Version        =   589884
         _ExtentX        =   6535
         _ExtentY        =   2514
         _StockProps     =   0
         ShowGroupBox    =   -1  'True
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H8000000C&
         Height          =   735
         Left            =   5040
         Top             =   720
         Width           =   405
      End
      Begin XtremeSuiteControls.ShortcutCaption sccTitle 
         Height          =   360
         Left            =   -60
         TabIndex        =   1
         Top             =   -30
         Width           =   7905
         _Version        =   589884
         _ExtentX        =   13944
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "基础设置>临床号源管理"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin ComctlLib.ImageList imgList16 
      Left            =   7380
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicSignalSourceManage.frx":1C02
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmClinicSignalSourceManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mlngModule As Long
Private mstrPrivs As String

Private Enum mRptHeadCol
    COL_ID = 0
    COL_排班方式图标
    COL_号类
    COL_号码
    COL_科室
    COL_收费项目
    COL_医生
    COL_建档
    COL_排班方式
    COL_预约天数
    COL_出诊频次
    COL_假日换休
    COL_假日控制状态
    COL_临床排班
    COL_适用性别
    COL_适用年龄段
    COL_是否停用
    COL_是否删除
    COL_建档时间
    COL_撤档时间
End Enum

Private mblnShowStopSignal As Boolean '是否显示已停用号源
Private Const conPane_SignalSorceList = 1
Private Const conPane_DetialList = 2
Private mrsWorkTime As ADODB.Recordset
Private mobj所有合作单位  As 合作单位控制集
Private mblnShowDetial As Boolean
Private mlngPreSel号源ID As Long
Private mintFindType As Integer
Private mrs号源 As Recordset

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区域控制
    '编制:刘兴洪
    '日期:2016-03-22 14:37:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String, panThis As Pane
    Dim panLeft As Pane
    
    Set panLeft = dkpMan.CreatePane(conPane_SignalSorceList, 200, 980, DockLeftOf, Nothing)
    panLeft.Title = "": panLeft.Tag = conPane_SignalSorceList
    panLeft.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panLeft.Handle = picBack.Hwnd
    Set panThis = dkpMan.CreatePane(conPane_DetialList, 100, 280, DockBottomOf, panLeft)
    panThis.Tag = conPane_DetialList
    panThis.Handle = picDetailList.Hwnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    'zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Sub

Private Sub LoadDetialData(ByVal lng号源Id As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载明细数据
    '入参:lng号源ID-号源ID
    '编制:刘兴洪
    '日期:2016-03-22 15:58:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj出诊记录集 As 出诊记录集
    Dim objPan As Pane
    On Error GoTo errHandle
    Set objPan = dkpMan.FindPane(conPane_DetialList)
    If Not mblnShowDetial Then
        If Not objPan Is Nothing Then
            If Not objPan.Closed Then objPan.Close
        End If
        If rptData.Visible Then rptData.SetFocus
        Exit Sub
    Else
       If Not objPan Is Nothing Then
            If Not objPan.Selected Then
                objPan.Select
            End If
        End If
    End If
    Screen.MousePointer = vbHourglass
    Set obj出诊记录集 = GetClinicRecordFromSignalSource(lng号源Id)
    Call CPDPages.LoadData(obj出诊记录集, Nothing, mobj所有合作单位, True)
    CPDPages.EditMode = ED_RegistPlan_View
    If rptData.Visible Then rptData.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
errHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo errHandler
    
    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '放在输出到Excel之后
        Set cbrControl = .Find(, conMenu_File_Excel)
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出为XML文件(&L)…", cbrControl.Index + 1)
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加号源(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改号源(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除号源(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用号源(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停用号源(&T)")
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowNoSourceDetial, "显示号源控制信息(&C)", cbrControl.index)
        cbrControl.Checked = mblnShowDetial
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "显示已停用号源(&S)", cbrControl.index)
        cbrControl.Checked = mblnShowStopSignal
        cbrControl.BeginGroup = True
    End With
    
    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加号源", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改号源", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除号源", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用号源", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停用号源", cbrControl.index + 1)
        .Item(cbrControl.index + 1).BeginGroup = True
    End With
    
    Set objPopup = cbrToolBar.Controls.Add(xtpControlButtonPopup, conMenu_View_FindType, "按号码过滤↓")
    objPopup.flags = xtpFlagRightAlign
    '被绑定的控件必须动态加载，因为工具栏一但被删除，被绑定的控件的句柄就会变成0
    Set objCustom = cbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Find, "")
    If txtFind.UBound > 0 Then Unload txtFind(1)
    Load txtFind(1)
    objCustom.Handle = txtFind(1).Hwnd
    objCustom.flags = xtpFlagRightAlign
    
    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("B"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '设置不常用命令
    '-----------------------------------------------------
    With mcbsMain.Options
'        .AddHiddenCommand conMenu_Edit_Archive
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnVisible As Boolean, blnEnable As Boolean
    Dim blnStop As Boolean '是否已停用
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            blnEnable = rptData.SelectedRows(0).Record(COL_是否删除).Value = ""
            blnStop = blnEnable And rptData.SelectedRows(0).Record(COL_是否停用).Value <> ""
        End If
    End If
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "出诊号源设置")

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = rptData.Rows.Count > 0
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnable And Not blnStop
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnable And Not blnStop
    Case conMenu_Edit_Reuse
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnStop
    Case conMenu_View_ShowNoSourceDetial
        Control.Checked = Control.Visible And mblnShowDetial
    Case conMenu_Edit_Stop
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And Not blnStop And blnEnable
    Case conMenu_View_FindType '查找方式
        Control.Caption = "按" & Decode(mintFindType, 0, "号码", 1, "科室", 2, "医生", "号码") & "过滤↓"
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '查找方式
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
    End Select
End Sub

Public Sub InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
        
    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "号码(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "科室(&2)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "医生(&3)"
            End If
        End With
    End Select
End Sub

Private Function ExcuteDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行删除操作
    '入参:lngID-号源ID
    '编制:刘兴洪
    '日期:2016-03-30 14:37:59
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim lngID As Long, str号码 As String
    On Error GoTo errHandle
    
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function

    lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
    str号码 = Trim(rptData.SelectedRows(0).Record(COL_号码).Caption)
    
    If Trim(rptData.SelectedRows(0).Record(COL_是否停用).Value) <> "" Then
        Call MsgBox("你要删除的号码为" & str号码 & "的号源已经被停用，不允许删除！", vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    If lngID = 0 Then
        MsgBox "当前未选中要删除的号源，不能进行删除操作！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

    '删除有效性检查
    If CheckIsUserPreRegist(str号码) Then
        If MsgBox("当前号源(号码为 " & str号码 & " )存在预约挂号记录，删除后，将会对该号源的所有出诊安排进行停诊，是否继续删除？", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        strSQL = "Select 1 From 临床出诊记录 Where 号源ID=[1] And 出诊日期+0>=Trunc(sysdate) And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        If Not rsTemp.EOF Then
            If MsgBox("当前号源(号码为 " & str号码 & " )存在有效出诊安排，删除后，将会对该号源的这些出诊安排进行停诊，是否继续删除？", _
                      vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("你确定要删除当前号源(号码为 " & str号码 & " )吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    strSQL = "Zl_临床出诊号源_Delete(" & lngID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    'rptData.Records (rptData.SelectedRows(0).Index)
    ExcuteDelete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim frmEdit As New frmClinicSignalSourceEdit, lngID As Long
    Dim str号码 As String
    Err = 0: On Error GoTo errHandler
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
            str号码 = Trim(rptData.SelectedRows(0).Record(COL_号码).Caption)
        End If
    End If
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem
        Dim strNewItem As String
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, Fun_Add, , strNewItem) Then Call LoadData(, strNewItem)
    Case conMenu_Edit_Modify
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, Fun_Update, lngID) Then Call LoadData
    Case conMenu_Edit_Delete
       If ExcuteDelete() Then Call LoadData
    Case conMenu_Edit_Reuse
        If StopAndResume(False) Then Call LoadData
    Case conMenu_Edit_Stop
        If StopAndResume(True) Then Call LoadData
    Case conMenu_View_ShowStoped '显示已停用号
        Control.Checked = Not Control.Checked
        mblnShowStopSignal = Control.Checked
        Call zlDatabase.SetPara("显示停用号源", IIf(mblnShowStopSignal, "1", "0"), glngSys, mlngModule)
        Call LoadData
    Case conMenu_View_ShowNoSourceDetial '显示明细信息
        mblnShowDetial = Not mblnShowDetial
        Control.Checked = mblnShowDetial
        Call zlDatabase.SetPara("显示缺省控制信息", IIf(mblnShowDetial, "1", "0"), glngSys, mlngModule)
        lngID = 0
        If rptData.SelectedRows.Count <> 0 Then
             If rptData.SelectedRows(0).GroupRow = False Then
                lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
             End If
        End If
        LoadDetialData (lngID)
        mlngPreSel号源ID = lngID
    Case conMenu_View_Refresh
        Call GetRecords: Call ExecuteFilter
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 3 '查找方式
        mintFindType = Val(Right(Control.ID, 2)) - 1
        mcbsMain.RecalcLayout
        txtFind(1).Text = ""
        If txtFind(1).Visible And txtFind(1).Enabled Then txtFind(1).SetFocus
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecuteFilter()
    '过滤数据
    Dim strKey As String
    
    Err = 0: On Error GoTo errHandler
    Call zlControl.TxtSelAll(txtFind(1))
    
    If Not mrs号源 Is Nothing Then
        With mrs号源
            If Trim(txtFind(1).Text) = "" Then
                .Filter = ""
            Else
                strKey = Replace(gstrLike, "%", "*") & UCase(txtFind(1).Text) & "*"
                Select Case mintFindType
                Case 0   '号码
                    .Filter = "号码 Like '" & strKey & "'"
                Case 1   '科室(简码)
                    .Filter = "科室 Like '" & strKey & "' Or 科室简码 Like '" & strKey & "'"
                Case 2   '医生(简码)
                    .Filter = "医生姓名 Like '" & strKey & "' Or 医生简码 Like '" & strKey & "'"
                Case Else
                    .Filter = ""
                End Select
            End If
        End With
    End If
    If mintFindType = 8 Then mintFindType = 0 '清除
    Call LoadData(False)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitGrid()
    Dim i As Long
    Dim objCol As ReportColumn, lngIdx As Long
    
    Err = 0: On Error GoTo errHandler
    With rptData
        .AutoColumnSizing = False '不使用自动列宽
        .AllowColumnRemove = False '不允许拖动删除列
        .ShowGroupBox = True '显示分组框
        .ShowItemsInGroups = False '不显示已分组的列
        .MultipleSelection = False '不允许多行选择
        .SetImageList Me.imgList16
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid '竖向表格线格式
            .HorizontalGridStyle = xtpGridSolid '横向表格线格式
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的内容..."
            .ShadeSortColor = .BackColor
            Set .CaptionFont = Me.Font
            Set .TextFont = Me.Font
        End With
    End With

    With rptData.Columns
        Set objCol = .Add(COL_ID, "ID", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_排班方式图标, "", 20, False)
        objCol.Groupable = False
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.AllowRemove = False
        
        Set objCol = .Add(COL_号类, "号类", 50, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_号码, "号码", 50, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_科室, "科室", 100, True)
        Set objCol = .Add(COL_收费项目, "收费项目", 120, True)
        Set objCol = .Add(COL_医生, "医生", 80, True)
        Set objCol = .Add(COL_建档, "建档", 50, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_排班方式, "排班方式", 55, True)
        Set objCol = .Add(COL_预约天数, "预约天数", 55, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_出诊频次, "出诊频次", 55, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_假日换休, "假日换休", 55, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_假日控制状态, "假日控制状态", 100, True)
        Set objCol = .Add(COL_临床排班, "临床排班", 55, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_适用性别, "适用性别", 55, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_适用年龄段, "适用年龄段", 70, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_是否停用, "是否停用", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_是否删除, "是否删除", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_建档时间, "建档时间", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_撤档时间, "撤档时间", 130, True): objCol.Alignment = xtpAlignmentCenter
    End With
    With rptData
    '        '将科室缺省升序排列
        .SortOrder.DeleteAll
        .SortOrder.Add .Columns(COL_号码)
        .SortOrder(0).SortAscending = True
        
        '将号类分组且升序排列
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns(COL_排班方式)
'        .GroupsOrder(0).SortAscending = True
        .Columns(COL_排班方式).Visible = False
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub GetRecords()
    '读取记录
    Dim strWhere As String, strSQL As String
    
    Err = 0: On Error GoTo errHandler
    Set mobj所有合作单位 = GetUnitsObjects(GetUnitAll())
    
'    If mblnShowDeleteSignal = False Then '不显示已删除
        strWhere = " And Nvl(a.是否删除,0) = 0"
'    End If
    If mblnShowStopSignal = False Then '不显示已停用
        strWhere = strWhere & _
            " And Nvl(a.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate" & vbNewLine & _
            " And Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate" & vbNewLine & _
            " And Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate" & vbNewLine & _
            " And Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate"
    End If
    
    '没有"所有科室"权限的操作员只能操作自己所属科室的号源
    If HavePrivs(mstrPrivs, "所有科室") = False Then
        strWhere = strWhere & "      And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = [1])"
    End If
    
    strSQL = "Select a.Id, a.号类, a.号码, c.名称 As 科室, c.简码 As 科室简码, b.名称 As 收费项目," & vbNewLine & _
            "        a.医生姓名, d.简码 As 医生简码,e.标识符, a.预约天数, a.出诊频次," & vbNewLine & _
            "        Nvl(a.是否建病案, 0) As 是否建病案,nvl(a.是否临床排班,0) as 是否临床排班," & vbNewLine & _
            "        Decode(nvl(a.假日控制状态,0), 1, '开放预约', 2, '禁止预约',3, '受节假日设置控制', '不上班') As 假日控制状态," & vbNewLine & _
            "        Decode(nvl(a.排班方式,0), 1, '按月排班', 2, '按周排班', '固定排班') As 排班方式," & vbNewLine & _
            "        Nvl(a.是否假日换休, 0) As 是否假日换休, Nvl(a.是否删除, 0) As 是否删除," & vbNewLine & _
            "        Case When Nvl(a.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate " & vbNewLine & _
            "               Or Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate " & vbNewLine & _
            "               Or Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate " & vbNewLine & _
            "               Or Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 1 Else 0 End As 是否停用," & vbNewLine & _
            "        a.建档时间, a.撤档时间, a.适用性别, a.适用年龄段" & vbNewLine & _
            " From 临床出诊号源 A, 收费项目目录 B, 部门表 C, 人员表 D,专业技术职务 E" & vbNewLine & _
            " Where a.项目id+0 = b.Id And a.科室id = c.Id(+) And a.医生ID = d.ID(+) and d.专业技术职务=e.名称(+)" & vbNewLine & _
            "        And Nvl(Nvl(c.站点,[3]),Nvl([2],'-')) = Nvl([2],'-')" & vbNewLine & _
                strWhere
    Set mrs号源 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, gstrNodeNo, gVisitPlan_ModulePara.str号源维护站点)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function LoadData(Optional ByVal blnReRead As Boolean = True, _
    Optional ByVal strNewItem As String) As Boolean
    '加载数据
    '入参：
    '   blnReRead 是否重新读取数据
    '   strNewItem 新增号源号码，用于定位
    Dim i As Long, j As Long, lngSelectRow As Long
    
    Err = 0: On Error GoTo errHandler
    Screen.MousePointer = vbHourglass

    If rptData.SelectedRows.Count > 0 Then lngSelectRow = rptData.SelectedRows(0).index
    rptData.Records.DeleteAll
    
    If mrs号源 Is Nothing Then
        Call GetRecords
    ElseIf mrs号源.State <> adStateOpen Then
        Call GetRecords
    ElseIf blnReRead Then
        Call GetRecords
    End If
    
    Do While Not mrs号源.EOF
        Call InsertRowData(Nvl(mrs号源!ID), Nvl(mrs号源!号类), Nvl(mrs号源!号码), Nvl(mrs号源!科室), _
            Nvl(mrs号源!收费项目), Nvl(mrs号源!医生姓名), Nvl(mrs号源!标识符), Nvl(mrs号源!是否建病案), Nvl(mrs号源!排班方式), _
            Nvl(mrs号源!预约天数), Nvl(mrs号源!适用性别), Nvl(mrs号源!适用年龄段), _
            Nvl(mrs号源!出诊频次), Nvl(mrs号源!是否假日换休), Nvl(mrs号源!假日控制状态), Nvl(mrs号源!是否临床排班), _
            Nvl(mrs号源!是否停用), Nvl(mrs号源!是否删除), Nvl(mrs号源!建档时间), Nvl(mrs号源!撤档时间))
        mrs号源.MoveNext
    Loop

    With rptData
        For i = 0 To .Records.Count - 1
            If i > .Records.Count - 1 Then Exit For
            If .Records(i).Item(COL_是否停用).Value <> "" _
                Or .Records(i).Item(COL_是否删除).Value <> "" Then
                For j = 0 To .Columns.Count - 1
                    .Records(i).Item(j).ForeColor = vbRed
                Next
            End If
        Next
    End With
    Call rptData.Populate '发布数据以更新界面
    If rptData.Rows.Count > 0 Then '该行选中且显示在可见区域
        If strNewItem <> "" Then
            For i = 0 To rptData.Rows.Count - 1
                If Not rptData.Rows(i).GroupRow Then
                    If rptData.Rows(i).Record(COL_号码).Caption = strNewItem Then
                        rptData.FocusedRow = rptData.Rows(i)
                        Exit For
                    End If
                End If
            Next
        Else
            If lngSelectRow = 0 Then
                rptData.FocusedRow = rptData.Rows(0)
            ElseIf lngSelectRow > rptData.Rows.Count - 1 Then
                rptData.FocusedRow = rptData.Rows(rptData.Rows.Count - 1)
            Else
                rptData.FocusedRow = rptData.Rows(lngSelectRow)
            End If
        End If
    End If
    Call SetReportControlBackColorAlternate(rptData)
    
    mlngPreSel号源ID = 0
    If rptData.SelectedRows.Count > 0 Then
         If rptData.SelectedRows(0).GroupRow = False Then
            mlngPreSel号源ID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
         End If
    End If
    Call LoadDetialData(mlngPreSel号源ID)
    
    Call mfrmMain.StatusShowInfoChanged(2, "当前共有" & mrs号源.RecordCount & "条号源信息")
    
    Screen.MousePointer = vbDefault
    Exit Function
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InsertRowData(ByVal strID As String, ByVal str号类 As String, ByVal str号码 As String, ByVal str科室 As String, _
    ByVal str项目 As String, ByVal str医生姓名 As String, ByVal str标识符 As String, ByVal str病案 As String, ByVal str排班方式 As String, _
    ByVal str预约天数 As String, ByVal str适用性别 As String, ByVal str适用年龄段 As String, _
    ByVal str出诊频次 As String, ByVal str假日换休 As String, ByVal str假日控制状态 As String, _
    ByVal str临床排班 As String, ByVal str是否停用 As String, ByVal str是否删除 As String, _
    ByVal str建档时间 As String, ByVal str撤档时间 As String)
    Dim objRecord As ReportRecord, ObjItem As ReportRecordItem
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    With rptData
        Set objRecord = .Records.Add()
        Set ObjItem = objRecord.AddItem(strID)
        Set ObjItem = objRecord.AddItem("")
        
        '图标设置
        Select Case str排班方式
        Case "按月排班"
            ObjItem.Icon = IIf(Val(str是否停用) = 0, 2, 3)
        Case "按周排班"
            ObjItem.Icon = IIf(Val(str是否停用) = 0, 4, 5)
        Case Else '固定排班
            ObjItem.Icon = IIf(Val(str是否停用) = 0, 0, 1)
        End Select
        
        Set ObjItem = objRecord.AddItem(str号类)
        Set ObjItem = objRecord.AddItem(str号码)
        If gVisitPlan_ModulePara.byt号码比较方式 = 1 Then
            ObjItem.Caption = str号码
            If Len(str号码) >= 5 Then '这样设置以后，号码就是按照数值排序了
                ObjItem.Value = str号码
            Else
                ObjItem.Value = String(5 - Len(str号码), "0") & str号码
            End If
        End If
        Set ObjItem = objRecord.AddItem(str科室)
        
        Set ObjItem = objRecord.AddItem(str项目)
        Set ObjItem = objRecord.AddItem(str医生姓名)
        ObjItem.Caption = str标识符 & str医生姓名
        Set ObjItem = objRecord.AddItem(IIf(Val(str病案) = 0, "", "√"))
        Set ObjItem = objRecord.AddItem(str排班方式)
        
        Set ObjItem = objRecord.AddItem(str预约天数)
        Set ObjItem = objRecord.AddItem(Val(str出诊频次))
        Set ObjItem = objRecord.AddItem(IIf(Val(str假日换休) = 0, "", "√"))
        Set ObjItem = objRecord.AddItem(str假日控制状态)
        Set ObjItem = objRecord.AddItem(IIf(Val(str临床排班) = 0, "", "√"))
        
        Set ObjItem = objRecord.AddItem(str适用性别)
        strTemp = str适用年龄段
        If InStr(strTemp, "~") > 0 Then
            If Split(strTemp, "~")(0) = "" Then
                strTemp = Split(strTemp, "~")(1) & "以下"
            ElseIf Split(strTemp, "~")(1) = "" Then
                strTemp = Split(strTemp, "~")(0) & "以上"
            End If
        End If
        Set ObjItem = objRecord.AddItem(strTemp)
        
        Set ObjItem = objRecord.AddItem(IIf(Val(str是否停用) = 0, "", "√"))
        Set ObjItem = objRecord.AddItem(IIf(Val(str是否删除) = 0, "", "√"))
        Set ObjItem = objRecord.AddItem(Format(str建档时间, "yyyy-mm-dd hh:mm:ss"))
        Set ObjItem = objRecord.AddItem(Format(str撤档时间, "yyyy-mm-dd hh:mm:ss"))
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.ActiveControl Is Nothing Then
        sccTitle.SetFocus
    ElseIf Not Me.ActiveControl Is txtFind(1) Then
        rptData.SetFocus
    End If
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    mlngPreSel号源ID = -1
    
    '读取参数值
    mblnShowDetial = Val(zlDatabase.GetPara("显示缺省控制信息", glngSys, mlngModule, "0")) = 1
    mblnShowStopSignal = Val(zlDatabase.GetPara("显示停用号源", glngSys, mlngModule, "0")) = 1
    Call InitPancel
    Call InitGrid
    
    RestoreWinState Me, App.ProductName
    Dim strFindType As String
    Call GetRegInFor(g私有模块, Me.Name, "FindType", strFindType)
    mintFindType = Val(strFindType)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    Call SaveRegInFor(g私有模块, Me.Name, "FindType", mintFindType)
    If Not mrs号源 Is Nothing Then Set mrs号源 = Nothing
    If Not mrsWorkTime Is Nothing Then Set mrsWorkTime = Nothing
End Sub

Private Sub picBack_Resize()
    Err = 0: On Error Resume Next
    
    With picBack
        shpBorder.Move 0, 0, .ScaleWidth - 6, .ScaleHeight - 6
        sccTitle.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        rptData.Left = .ScaleLeft + 10
        rptData.Top = sccTitle.Top + sccTitle.Height
        rptData.Width = .ScaleWidth - 30
        rptData.Height = .ScaleHeight - sccTitle.Height - 30
    End With
End Sub
 
Private Sub picDetailList_Resize()
    Err = 0: On Error Resume Next
    With picDetailList
        CPDPages.Left = .ScaleLeft
        CPDPages.Top = .ScaleTop
        CPDPages.Width = .ScaleWidth
        CPDPages.Height = .ScaleHeight
    End With
End Sub

Private Sub rptData_ColumnOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub

Private Sub rptData_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim frmEdit As New frmClinicSignalSourceEdit, lngID As Long
    Dim blnStop As Boolean
    
    Err = 0: On Error GoTo errHandler
    If rptData.SelectedRows.Count = 0 Then Exit Sub
    If rptData.SelectedRows(0).GroupRow Then Exit Sub
    
    lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
    blnStop = rptData.SelectedRows(0).Record(COL_是否停用).Value <> ""
    If zlStr.IsHavePrivs(mstrPrivs, "出诊号源设置") And blnStop = False Then
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, Fun_Update, lngID) Then Call LoadData '刷新数据
    Else
        frmEdit.ShowMe Me, mlngModule, mstrPrivs, Fun_View, lngID
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_SelectionChanged()
    Dim lng号源Id As String
    
    Err = 0: On Error GoTo errHandler
    lng号源Id = 0
    If rptData.SelectedRows.Count <> 0 Then
        With rptData.SelectedRows(0)
            If Not .GroupRow Then
                lng号源Id = Val(.Record(COL_ID).Value)
            End If
        End With
    End If
    If mlngPreSel号源ID = lng号源Id Then Exit Sub
    
    mlngPreSel号源ID = lng号源Id
    Call LoadDetialData(lng号源Id)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptData_SortOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub

Private Function StopAndResume(ByVal blnStop As Boolean) As Boolean
    '功能：停用或启用号源
    '返回：停用或启用成功,返回true,否则返回False
    Dim i As Integer, intRow As Integer
    Dim strSQL As String, str号码 As String, lng号源Id As Long
    Dim rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If rptData.SelectedRows.Count = 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    str号码 = Trim(rptData.SelectedRows(0).Record(COL_号码).Caption)
    lng号源Id = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
    
    If blnStop Then
        If CheckIsUserPreRegist(str号码) Then
            If MsgBox("号码为" & str号码 & "的号源已经存在预约挂号记录。停用该号源后，将会对该号源的所有出诊安排进行停诊，是否继续停用？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("停用该号源后，将会对该号源的所有出诊安排进行停诊，是否继续停用号码为" & str号码 & "的号源？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    Else
        If MsgBox("你确认要" & IIf(blnStop, "停用", "启用") & "号码为""" & str号码 & """的号源吗？", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        strSQL = "Select Case When Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 1 Else 0 End As 部门停用," & vbNewLine & _
                "        Case When Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 1 Else 0 End As 人员停用," & vbNewLine & _
                "        Case When Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 1 Else 0 End As 项目停用" & vbNewLine & _
                " From 临床出诊号源 A, 部门表 B, 人员表 C, 收费项目目录 D" & vbNewLine & _
                " Where a.科室id = b.Id And a.医生id = c.Id(+) And a.项目ID = d.ID And a.Id = [1]" & vbNewLine & _
                "       And (Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate" & vbNewLine & _
                "            Or Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate" & vbNewLine & _
                "            Or Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng号源Id)
        If Not rsTemp.EOF Then
            If Val(Nvl(rsTemp!人员停用)) = 1 Then
                MsgBox "该号源的医生已被停用或删除，在未重新启用医生前不能启用该号源！", vbInformation, gstrSysName
                Exit Function
            ElseIf Val(Nvl(rsTemp!部门停用)) = 1 Then
                MsgBox "该号源的科室已被停用或删除，在未重新启用科室前不能启用该号源！", vbInformation, gstrSysName
                Exit Function
            ElseIf Val(Nvl(rsTemp!项目停用)) = 1 Then
                MsgBox "该号源的收费项目已被停用或删除，在未重新启用收费项目前不能启用该号源！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    strSQL = "Zl_临床出诊号源_Stopandstart(" & lng号源Id & "," & IIf(blnStop, 1, 0) & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    StopAndResume = True
    
    If blnStop = False Then '启用时重新生成没有生成的出诊记录
        '重新生成出诊记录
        strSQL = "Zl1_Auto_Buildingregisterplan(Null)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckIsUserPreRegist(ByVal str号码 As String) As Boolean
    '功能:检查是否存在预约挂号
    '返回:存在返回true,否则返回False
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    strSQL = "Select 1 From 门诊费用记录" & vbNewLine & _
            " Where 记录性质=4 And 记录状态=0 And 发生时间>=Sysdate And 计算单位=[1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str号码)
    If Not rsTemp.EOF Then
        CheckIsUserPreRegist = True
        Exit Function
    End If

    strSQL = "Select 1 From 病人挂号记录" & vbNewLine & _
            " Where 记录性质=1 And 记录状态=1 And 发生时间>=Sysdate And 号别=[1] And Rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str号码)
    If Not rsTemp.EOF Then
        CheckIsUserPreRegist = True
        Exit Function
    End If
    CheckIsUserPreRegist = False
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub zlDataPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If UserInfo.姓名 = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte, strHiddenCols As String
    
    Err = 0: On Error GoTo errHandler
    objOut.Title.Text = "临床号源清单"
    '将ReportControl转换为VSFlexGrid
    strHiddenCols = CStr(COL_ID) & "," & CStr(COL_排班方式图标) & "," & _
        CStr(COL_是否删除) & "," & CStr(COL_是否停用)
    Set objOut.Body = GetVsfGridData(rptData, strHiddenCols)
    
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If rptData.Visible Then rptData.SetFocus
End Sub

Private Sub txtFind_KeyPress(index As Integer, KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        Call ExecuteFilter
    End If
End Sub

Private Sub txtFind_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then
        '按了右键菜单快捷键，清除粘贴板内容
        If Clipboard.GetText <> "" Then Clipboard.Clear
    End If
End Sub

Private Sub txtFind_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtFind(index).Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtFind(index).Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtFind_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtFind(index).Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
