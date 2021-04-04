VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmClinicPlanStopVisitManage 
   Caption         =   "停诊安排管理"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9645
   Icon            =   "frmClinicPlanStopVisitManage.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9645
   StartUpPosition =   2  '屏幕中心
   Begin XtremeReportControl.ReportControl rptData 
      Height          =   2505
      Left            =   1830
      TabIndex        =   0
      Top             =   1200
      Width           =   5385
      _Version        =   589884
      _ExtentX        =   9499
      _ExtentY        =   4419
      _StockProps     =   0
      ShowGroupBox    =   -1  'True
   End
   Begin VB.PictureBox picButton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9645
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   5940
      Visible         =   0   'False
      Width           =   9645
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   450
         TabIndex        =   4
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "退出(&E)"
         Height          =   350
         Left            =   7830
         TabIndex        =   3
         Top             =   180
         Width           =   1100
      End
   End
   Begin ComctlLib.ImageList imgList16 
      Left            =   6690
      Top             =   4500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicPlanStopVisitManage.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicPlanStopVisitManage.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicPlanStopVisitManage.frx":13FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClinicPlanStopVisitManage.frx":1998
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   1350
      TabIndex        =   1
      Top             =   480
      Width           =   7905
      _Version        =   589884
      _ExtentX        =   13944
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "出诊安排>停诊安排"
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
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   495
      Left            =   240
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmClinicPlanStopVisitManage"
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
    Col_记录ID
    COL_图标
    Col_状态
    COL_申请人
    COL_停诊号码
    COL_开始时间
    COL_终止时间
    COL_停诊原因
    COL_申请时间
    COL_审批人
    COL_审批时间
    COL_失效时间
    COL_登记人
End Enum
Private mstrFilter As String
Private mstrDefaultFilter  As String
Private Type Type_SQLCondition
    ApplyName As String
    AuditName As String
    StopBegin As Date
    StopEnd As Date
End Type
Private SQLCondition As Type_SQLCondition

Private mstrDoctorName As String
Private mblnShowDoctorStopVisit As Boolean

Public Sub ShowDoctorStopVisit(frmParent As Form, ByVal strDoctorName As String)
    '显示指定医生停诊信息
    mstrDoctorName = strDoctorName
    
    If strDoctorName = "" Then Exit Sub
    On Error Resume Next
    mblnShowDoctorStopVisit = True
    Me.Show 1, frmParent
End Sub

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddMonthPlan, "制定月出诊表(&Y)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddWeekPlan, "制定周出诊表(&W)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "停诊申请(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "取消申请(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "停诊审批(&V)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnAudit, "取消审批(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "终止安排(&C)"): cbrControl.BeginGroup = True
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤(&F)", cbrControl.index)
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddMonthPlan, "月出诊表", cbrControl.index + 1): cbrControl.BeginGroup = True
        cbrControl.ToolTipText = "制定月出诊表"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddWeekPlan, "周出诊表", cbrControl.index + 1)
        cbrControl.ToolTipText = "制定周出诊表"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "停诊申请", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "取消申请", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "停诊审批", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnAudit, "取消审批", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "终止安排", cbrControl.index + 1): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤", cbrControl.index + 1): cbrControl.BeginGroup = True
    End With
    
    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("Y"), conMenu_Edit_AddMonthPlan
        .Add FCONTROL, Asc("W"), conMenu_Edit_AddWeekPlan
        
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
        .Add FCONTROL, Asc("V"), conMenu_Edit_Audit
        .Add FCONTROL, Asc("C"), conMenu_Edit_UnAudit
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
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
    Dim blnEnable As Boolean, blnAudit As Boolean, blnStop As Boolean
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    If Control.ID = conMenu_Edit_Delete Or Control.ID = conMenu_Edit_Audit _
        Or Control.ID = conMenu_Edit_UnAudit Or Control.ID = conMenu_Edit_Stop Then
        If rptData.SelectedRows.Count > 0 Then
            If Not rptData.SelectedRows(0).GroupRow Then
                blnAudit = rptData.SelectedRows(0).Record(COL_审批人).Value <> ""
                blnEnable = Val(rptData.SelectedRows(0).Record(Col_记录ID).Value) = 0 '有出诊记录的不允许任何操作
                blnStop = rptData.SelectedRows(0).Record(COL_失效时间).Value <> ""
            End If
        End If
    End If

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = rptData.Rows.Count > 0
    Case conMenu_EditPopup
        Control.Visible = (mfrmMain.mFunListActived And (HavePrivs(mstrPrivs, "出诊安排"))) _
            Or (mfrmMain.mFunListActived = False And (HavePrivs(mstrPrivs, "出诊安排;停诊申请;停诊审批")))
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AddMonthPlan '制定月出诊表
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AddWeekPlan '制定周出诊表
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem '停诊申请
        Control.Visible = HavePrivs(mstrPrivs, "停诊申请")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Delete '取消申请
        Control.Visible = HavePrivs(mstrPrivs, "停诊申请")
        Control.Enabled = Control.Visible And blnEnable And Not blnAudit
    Case conMenu_Edit_Audit '停诊审批
        Control.Visible = HavePrivs(mstrPrivs, "停诊审批")
        Control.Enabled = Control.Visible And blnEnable And Not blnAudit
    Case conMenu_Edit_UnAudit '取消审批
        Control.Visible = HavePrivs(mstrPrivs, "停诊审批")
        Control.Enabled = Control.Visible And blnEnable And blnAudit And Not blnStop
    Case conMenu_Edit_Stop '终止安排
        Control.Visible = HavePrivs(mstrPrivs, "停诊审批")
        Control.Enabled = Control.Visible And blnEnable And blnAudit And Not blnStop
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim frmEdit As New frmClinicPlanStopVisitEdit, lngID As Long
    Dim strApplyName As String
    
    Err = 0: On Error GoTo errHandler
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
        End If
    End If
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem '停诊申请
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, 1, , strApplyName) Then Call LoadData(strApplyName)
    Case conMenu_Edit_Delete '取消申请
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, 2, lngID) Then Call LoadData
    Case conMenu_Edit_Audit '停诊审批
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, 3, lngID) Then Call LoadData
    Case conMenu_Edit_UnAudit '取消审批
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, 4, lngID) Then Call LoadData
    Case conMenu_Edit_Stop '终止安排
        If frmEdit.ShowMe(Me, mlngModule, mstrPrivs, 5, lngID) Then Call LoadData
    Case conMenu_View_Refresh
        Call LoadData '刷新数据
    Case conMenu_View_Filter '过滤
        With frmClinicPlanStopVisitFilter
            .mblnOk = False
            .Show 1, Me
            If .mblnOk Then
                mstrFilter = .mstrFilter
                SQLCondition.ApplyName = Trim(.txtApply.Text)
                SQLCondition.AuditName = Trim(.txtAudit.Text)
                SQLCondition.StopBegin = .dtpStopBegin.Value
                SQLCondition.StopEnd = .dtpStopEnd.Value
                Call LoadData
            End If
        End With
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub RefreshData()
    Err = 0: On Error GoTo errHandler
    
    Call LoadData
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnShowDoctorStopVisit Then Exit Sub
    If Me.ActiveControl Is Nothing Then
        sccTitle.SetFocus
    Else
        rptData.SetFocus
    End If
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    
    Call InitGrid
    mstrFilter = " And Nvl(失效时间,终止时间)>sysdate"
    
    If mblnShowDoctorStopVisit Then
        If LoadData() = False Then
            MsgBox "医生 " & mstrDoctorName & " 当前无有效停诊安排！", vbInformation + vbOKOnly, gstrSysName
            Unload Me: Exit Sub
        End If
        shpBorder.Visible = False
        sccTitle.Visible = False
        Me.Caption = mstrDoctorName & " 停诊安排"
        
        picButton.Visible = True
    End If
    RestoreWinState Me, App.ProductName
    '是否显示分组框
    rptData.ShowGroupBox = (mblnShowDoctorStopVisit = False)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If mblnShowDoctorStopVisit = False Then
        shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
        sccTitle.Move 8, 8, shpBorder.Width - 20
    End If
    
    With rptData
        .Left = 8
        .Top = IIf(mblnShowDoctorStopVisit, 0, sccTitle.Top + sccTitle.Height)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - IIf(mblnShowDoctorStopVisit, picButton.Height, 0) - .Top - 20
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrFilter = ""
    mblnShowDoctorStopVisit = False
    SaveWinState Me, App.ProductName
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
            Set .PreviewTextFont = Me.Font
        End With
    End With
    
    With rptData.Columns
        Set objCol = .Add(COL_ID, "ID", 50, True): objCol.Visible = False
        Set objCol = .Add(Col_记录ID, "记录ID", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_图标, "", 20, False)
        objCol.Groupable = False
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.AllowRemove = False
        Set objCol = .Add(Col_状态, "状态", 50, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_申请人, "申请人", 50, True)
        Set objCol = .Add(COL_停诊号码, "停诊号码", 100, True)
        Set objCol = .Add(COL_开始时间, "开始时间", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_终止时间, "终止时间", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_停诊原因, "停诊原因", 140, True)
        Set objCol = .Add(COL_申请时间, "申请时间", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_审批人, "审批人", 50, True)
        Set objCol = .Add(COL_审批时间, "审批时间", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_失效时间, "失效时间", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_登记人, "登记人", 50, True)
    End With
    
    With rptData
        '将申请人分组显示
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns(COL_申请人)
        .Columns(COL_申请人).Visible = False
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadData(Optional ByVal strApplyName As String) As Boolean
    '入参：
    '   strApplyName - 缺省定位到该医生
    Dim i As Long, j As Long
    Dim lngSelectRow As Long
    Dim strSQL As String, rsData As ADODB.Recordset
    Dim objRecord As ReportRecord, ObjItem As ReportRecordItem
    Dim dtNow As Date
    
    Err = 0: On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    If rptData.SelectedRows.Count > 0 Then lngSelectRow = rptData.SelectedRows(0).index
    rptData.Records.DeleteAll
    
    If mblnShowDoctorStopVisit Then mstrFilter = " And Nvl(失效时间,终止时间)>sysdate And 申请人=[1]"
    strSQL = "Select ID, 记录ID, 停诊原因, 开始时间, 终止时间, 申请人, 申请时间, 审批人, 审批时间," & vbNewLine & _
            "       Decode(审批时间,NULL,'未审批','已审批') As 状态, 登记人, 失效时间,Nvl(停诊号码,'-') As 停诊号码" & vbNewLine & _
            " From 临床出诊停诊记录" & vbNewLine & _
            " Where 记录ID Is Null " & mstrFilter
    If mblnShowDoctorStopVisit = False Then
        With SQLCondition
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .ApplyName, .AuditName, .StopBegin, .StopEnd)
        End With
    Else
        strSQL = strSQL & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select ID, 记录id, 停诊原因, 开始时间, 终止时间, 申请人, 申请时间, 审批人, 审批时间," & vbNewLine & _
            "        Decode(审批时间, Null, '未审批', '已审批') As 状态, 登记人, 失效时间,Nvl(停诊号码,'-') As 停诊号码" & vbNewLine & _
            " From 临床出诊停诊记录 A" & vbNewLine & _
            " Where 记录id Is Not Null " & mstrFilter & vbNewLine & _
            "       And Not Exists(Select 1" & vbNewLine & _
            "                       From 临床出诊停诊记录" & vbNewLine & _
            "                       Where 记录id Is Null " & mstrFilter & " And a.开始时间 >= 开始时间 And Nvl(a.失效时间,a.终止时间) <= Nvl(失效时间,终止时间))"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrDoctorName)
    End If
    If rsData.RecordCount = 0 Then
        Screen.MousePointer = vbDefault
        Call rptData.Populate '发布数据以更新界面
        Exit Function
    End If
    
    dtNow = zlDatabase.Currentdate
    Do While Not rsData.EOF
        Set objRecord = rptData.Records.Add()
        objRecord.AddItem Nvl(rsData!ID)
        objRecord.AddItem Nvl(rsData!记录ID)
        '图标设置
        Set ObjItem = objRecord.AddItem("")
        If Nvl(rsData!审批时间) = "" Then '未审批
            ObjItem.Icon = IIf(CDate(Nvl(rsData!失效时间, rsData!终止时间)) > dtNow, 0, 1)
        Else '已审批
            ObjItem.Icon = IIf(CDate(Nvl(rsData!失效时间, rsData!终止时间)) > dtNow, 2, 3)
        End If
        objRecord.AddItem Nvl(rsData!状态)
        objRecord.AddItem Nvl(rsData!申请人)
        objRecord.AddItem Nvl(rsData!停诊号码)
        objRecord.AddItem Format(Nvl(rsData!开始时间), "yyyy-mm-dd hh:mm:ss")
        objRecord.AddItem Format(Nvl(rsData!终止时间), "yyyy-mm-dd hh:mm:ss")
        objRecord.AddItem Nvl(rsData!停诊原因)
        objRecord.AddItem Format(Nvl(rsData!申请时间), "yyyy-mm-dd hh:mm:ss")
        objRecord.AddItem Nvl(rsData!审批人)
        objRecord.AddItem Format(Nvl(rsData!审批时间), "yyyy-mm-dd hh:mm:ss")
        objRecord.AddItem Format(Nvl(rsData!失效时间), "yyyy-mm-dd hh:mm:ss")
        objRecord.AddItem Format(Nvl(rsData!登记人), "yyyy-mm-dd hh:mm:ss")
        rsData.MoveNext
    Loop
    Call rptData.Populate '发布数据以更新界面
    '已审批用蓝色显示
    For i = 0 To rptData.Records.Count - 1
        If rptData.Records(i).Item(COL_审批时间).Value <> "" Then
            For j = 0 To rptData.Columns.Count - 1
                rptData.Records(i).Item(j).ForeColor = vbBlue
            Next
        End If
    Next
    
    If rptData.Rows.Count > 0 Then '该行选中且显示在可见区域
        If strApplyName <> "" Then
            For i = 0 To rptData.Rows.Count - 1
                If Not rptData.Rows(i).GroupRow Then
                    If rptData.Rows(i).Record(COL_申请人).Value = strApplyName Then
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
    Screen.MousePointer = vbDefault
    LoadData = True
    Exit Function
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub picButton_Resize()
    On Error Resume Next
    cmdExit.Left = picButton.ScaleWidth - cmdExit.Width - 500
End Sub

Private Sub rptData_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo errHandler
    If mblnShowDoctorStopVisit Then Exit Sub
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

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If rptData.Visible Then rptData.SetFocus
End Sub

Private Sub zlDataPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If UserInfo.姓名 = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte, strHiddenCols As String
    
    Err = 0: On Error GoTo errHandler
    objOut.Title.Text = "停诊安排清单"
    '将ReportControl转换为VSFlexGrid
    strHiddenCols = CStr(COL_ID) & "," & CStr(Col_记录ID) & "," & CStr(COL_图标)
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

