VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmPatinetAuditing 
   Caption         =   "病人审核"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   Icon            =   "frmPatinetAuditing.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPatientList 
      Height          =   5265
      Left            =   360
      ScaleHeight     =   5205
      ScaleWidth      =   4065
      TabIndex        =   0
      Top             =   870
      Width           =   4125
      Begin XtremeReportControl.ReportControl rptPatientList 
         Height          =   3945
         Left            =   750
         TabIndex        =   1
         Top             =   780
         Width           =   2415
         _Version        =   589884
         _ExtentX        =   4260
         _ExtentY        =   6959
         _StockProps     =   0
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cbo时间 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   4890
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker DTPDate 
         Height          =   300
         Left            =   1470
         TabIndex        =   2
         Top             =   4890
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   149815297
         CurrentDate     =   40049
      End
      Begin MSComCtl2.DTPicker dtpDateEnd 
         Height          =   300
         Left            =   2790
         TabIndex        =   4
         Top             =   4890
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   149815297
         CurrentDate     =   40049
      End
   End
   Begin MSComctlLib.ImageList Imglist 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":6852
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":6DEC
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":7386
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":7920
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":7EBA
            Key             =   ""
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":8454
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":87EE
            Key             =   ""
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":8B88
            Key             =   ""
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":8F22
            Key             =   ""
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":92BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":FB1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":16380
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":1CBE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":23444
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":29CA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   1980
      Top             =   270
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPatinetAuditing.frx":30508
      Left            =   1260
      Top             =   270
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPatinetAuditing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mfrmWrite As frmLisStationWrite                  '报告填写窗体
Attribute mfrmWrite.VB_VarHelpID = -1
Private mlngPatienID As Long                                        '病人ID
Private mstrPrivs   As String                                       '权限
Private mstrAuditingMan As String                                   '审核人,审核时传入过程
Private Enum mCol
    病人ID
    姓名
    性别
    年龄
    床号
    科室
    来源
    主页ID
    标识号
    出生日期
    单位
End Enum

Private Sub CreateCbs()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Set Me.cbrthis.Icons = zlCommFun.GetPubIcons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False
    

    '-----------------------------------------------------
    '菜单定义
    Me.cbrthis.ActiveMenuBar.Title = "菜单"
'    Me.cbrthis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&T)…"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告预览(&V)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With
    
    
    'conMenu_EditPopup
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "报告(&E)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "报告审核(&A)")
    End With



'    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
            cbrPopControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&F)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With

   

    '快键绑定
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_Edit_Audit
        
        
    End With

    '设置不常用菜单
'    With Me.cbrthis.Options
'        .AddHiddenCommand conMenu_File_PrintSet
'        .AddHiddenCommand conMenu_File_Excel
'        .AddHiddenCommand conMenu_View_Jump
'        .AddHiddenCommand conMenu_View_Refresh
'    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbrthis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审报告"): cbrControl.BeginGroup = True
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    With cbo时间
        .AddItem "今  天"
        .AddItem "昨  天"
        .AddItem "本  周"
        .AddItem "本  月"
        .AddItem "本  季"
        .AddItem "本半年"
        .AddItem "本  年"
        .AddItem "前三天"
        .AddItem "前一周"
        .AddItem "前半月"
        .AddItem "前一月"
        .AddItem "前二月"
        .AddItem "前三月"
        .AddItem "前半年"
        .AddItem "自定义"
    End With
    cbo时间.Text = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0)
End Sub
Private Sub CreateDockPane()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    Dim lngPane5Width As Long, lngPane2Height As Long, lngPane2Width As Long, lngPane3Height As Long
    
    Set mfrmWrite = New frmLisStationWrite                          '报告填写窗体
    mfrmWrite.mblnPatientFind = True
    mfrmWrite.fraComment.Tag = "不显示"
    dkpMain.Options.HideClient = True
    
    Set Pane1 = dkpMain.CreatePane(1, 150, 150, DockLeftOf, Nothing)
    Pane1.Title = "病人列表"
    Pane1.Handle = Me.picPatientList.hWnd
    Pane1.Options = PaneNoHideable Or PaneNoCloseable Or PaneNoFloatable

    Set Pane2 = dkpMain.CreatePane(2, 400, 600, DockRightOf, Nothing)
    Pane2.Title = "结果信息"
    Pane2.Handle = mfrmWrite.hWnd
    Pane2.Options = PaneNoHideable Or PaneNoCloseable Or PaneNoFloatable
    
    Pane1.Select
    mfrmWrite.fraComment.Tag = "不显示"
    mfrmWrite.mblnPatientFind = True
End Sub

Private Sub cbo时间_Click()
    zlDatabase.SetPara "标本范围", cbo时间.Text & ";" & Me.DTPDate & ";" & Me.dtpDateEnd, 100, 1208
    Me.DTPDate.Visible = (Me.cbo时间.Text = "自定义")
    Me.dtpDateEnd.Visible = (Me.cbo时间.Text = "自定义")
    '刷新
    If Me.Visible = True Then
        Call RefreshDate
    End If
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
        '------------------------------------------------------------------------------------------
        Case conMenu_File_PrintSet                                              '打印设置
        Case conMenu_File_Preview                                               '预览
        Case conMenu_File_Print                                                 '打印
        Case conMenu_File_Exit                                                  '退出
            Unload Me
        '------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Button                                                '标准按钮
            Control.Checked = Not Control.Checked
            Me.cbrthis(2).Visible = Control.Checked
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text                                                  '文本标签
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbrthis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_ToolBar_Size                                                  '大图标
            Control.Checked = Not Control.Checked
            Me.cbrthis.Options.LargeIcons = Not Me.cbrthis.Options.LargeIcons
            Me.cbrthis.RecalcLayout
        
        
        '------------------------------------------------------------------------------------------
        Case conMenu_Edit_Audit                                                 '审核
            If Not Me.rptPatientList.FocusedRow Is Nothing Then
                Call AuditingPatient(Me.rptPatientList.FocusedRow.Record(mCol.病人ID).Value)
                Call RefreshDate
            End If
            
        '------------------------------------------------------------------------------------------
        Case conMenu_View_Refresh                                               '刷新
            RefreshDate
        '------------------------------------------------------------------------------------------
        Case conMenu_Help_Help                                                          '帮助主题
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_Web                                                           'WEB上的
            Call zlHomePage(hWnd)
        
        Case conMenu_Help_Web_Home                                                      '主页
            Call zlHomePage(Me.hWnd)
        
        Case conMenu_Help_Web_Mail                                                      '发送反馈
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_Help_About                                                         '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button                                                    '显示工具条
            Control.Checked = Me.cbrthis(2).Visible
        Case conMenu_View_ToolBar_Text                                                      '是否显示文字
            Control.Checked = Not (Me.cbrthis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size                                                      '是否显示大图标
            Control.Checked = Me.cbrthis.Options.LargeIcons
    End Select
End Sub

Private Sub dkpMain_Resize()
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    Me.cbrthis.RecalcLayout
End Sub

Private Sub dtpDate_Change()
    zlDatabase.SetPara "标本范围", cbo时间.Text & ";" & Me.DTPDate & ";" & Me.dtpDateEnd, 100, 1208
    Call RefreshDate
End Sub

Private Sub dtpDateEnd_Change()
    zlDatabase.SetPara "标本范围", cbo时间.Text & ";" & Me.DTPDate & ";" & Me.dtpDateEnd, 100, 1208
    Call RefreshDate
End Sub

Private Sub Form_Load()
    Call CreateCbs
    Call CreateDockPane
    Call CreaterptListHead
    
    DTPDate.Value = Now
    dtpDateEnd.Value = Now
    
    
    Call RefreshDate
    
    Call RestoreWinState(Me, App.ProductName)                   '界面恢复
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.dkpMain.RecalcLayout
    Call mfrmWrite.zlRefreshPatient(mlngPatienID)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Unload mfrmWrite
End Sub

Private Sub picPatientList_Resize()
    On Error Resume Next
    Me.rptPatientList.Top = 10
    Me.rptPatientList.Left = 10
    Me.rptPatientList.Width = Me.picPatientList.ScaleWidth - 20
    Me.rptPatientList.Height = Me.picPatientList.ScaleHeight - Me.cbo时间.Height - 40
    
    With Me.cbo时间
        .Top = Me.picPatientList.ScaleHeight - .Height - 20
        .Left = 0
    End With
    With Me.DTPDate
        .Top = Me.cbo时间.Top
        .Left = Me.cbo时间.Left + Me.cbo时间.Width + 20
    End With
    With Me.dtpDateEnd
        .Top = Me.cbo时间.Top
        .Left = Me.DTPDate.Left + Me.DTPDate.Width + 20
    End With
End Sub
Private Sub CreaterptListHead()
    Dim Column As ReportColumn
    Dim i As Integer
    With Me.rptPatientList.Columns
        
        rptPatientList.AllowColumnRemove = False
        rptPatientList.ShowItemsInGroups = False
        rptPatientList.SetImageList Imglist
        With rptPatientList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
        Set Column = .Add(mCol.病人ID, "病人ID", 45, False): Column.Visible = False
        Set Column = .Add(mCol.姓名, "姓名", 55, True)
        Set Column = .Add(mCol.性别, "性别", 40, True)
        Set Column = .Add(mCol.年龄, "年龄", 40, True)
        Set Column = .Add(mCol.来源, "来源", 40, True)
        Set Column = .Add(mCol.标识号, "标识号", 55, True)
        Set Column = .Add(mCol.床号, "床号", 40, True)
        Set Column = .Add(mCol.科室, "科室", 75, True)
        Set Column = .Add(mCol.主页ID, "主页ID", 55, True): Column.Visible = False
        Set Column = .Add(mCol.出生日期, "出生日期", 75, True)
        Set Column = .Add(mCol.单位, "单位", 120, True)
        
    End With
End Sub
Private Sub RefreshDate()
    '刷新病人列表数据
    Dim rsTmp As New adodb.Recordset
    Dim strSQL As String
    Dim strStart As String
    Dim strEnd As String
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim intIndex As Integer
    Dim blnSelect As Boolean
    
    On Error GoTo errH
    
    strStart = GetDateTime(Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0), 2)
    
    If strStart = "自定义" Then
        strStart = Format(Me.DTPDate, "yyyy-mm-dd 00:00:00")
        strEnd = Format(Me.dtpDateEnd, "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    End If
    
    
    strSQL = "Select distinct 病人id, 姓名, 性别, 年龄, decode(病人来源,1,'门诊',2,'住院',3,'外来',4,'体验','其他') as 病人来源, 标识号, 床号,b.名称 as 病人科室, 主页id, 出生日期," & vbNewLine & _
            "Decode(a.样本状态, 1, '检验中', 2, '已检验') As 执行状态,decode(p.项目,'任务团体',p.内容,'') as 单位 " & vbNewLine & _
            "From 检验标本记录 a,部门表 b,病人医嘱附件 P " & vbNewLine & _
            "Where 核收时间 between [1] And [2] " & vbNewLine & _
            " and a.申请科室ID =  B.ID(+) and a.医嘱ID is not null And 审核人 is null and 微生物标本 is null and a.医嘱id = P.医嘱ID(+) "
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strStart), CDate(strEnd))
    
    If Not Me.rptPatientList.FocusedRow Is Nothing Then intIndex = Me.rptPatientList.FocusedRow.Index
    
    Me.rptPatientList.Records.DeleteAll
    mfrmWrite.zlRefreshPatient (-1)
    
    Do While Not rsTmp.EOF
        With Me.rptPatientList
            Set Record = .Records.Add
            For intLoop = 0 To .Columns.Count - 1
                Record.AddItem ""
            Next
'            If rsTmp("执行状态") = "已检验" Then
'                Record.Item(mCol.执行状态).Value = "已检验"
'                Record.Item(mCol.执行状态).Icon = 7
'            End If
            Record(mCol.病人ID).Value = rsTmp("病人ID")
            Record(mCol.姓名).Value = Nvl(rsTmp("姓名"))
            Record(mCol.年龄).Value = Nvl(rsTmp("年龄"))
            Record(mCol.性别).Value = Nvl(rsTmp("性别"))
            Record(mCol.床号).Value = Nvl(rsTmp("床号"))
            Record(mCol.来源).Value = Nvl(rsTmp("病人来源"))
            Record(mCol.标识号).Value = Nvl(rsTmp("标识号"))
            Record(mCol.科室).Value = Nvl(rsTmp("病人科室"))
            Record(mCol.主页ID).Value = Nvl(rsTmp("主页id"))
            Record(mCol.出生日期).Value = Nvl(rsTmp("出生日期"))
            
'            If Nvl(rsTmp("项目")) = "任务团体" Then
                Record.Item(mCol.单位).Value = Nvl(rsTmp("单位"))
'            End If
            
            If mlngPatienID = rsTmp("病人ID") Then
                blnSelect = True
                intIndex = Record.Index
            End If
        End With
        rsTmp.MoveNext
    Loop
    
    Me.rptPatientList.Populate
    
    If Me.rptPatientList.Rows.Count > 0 Then
        If blnSelect = True Then
            Me.rptPatientList.FocusedRow = Me.rptPatientList.Rows(intIndex)
        Else
            If intIndex >= Me.rptPatientList.Rows.Count Then
                Me.rptPatientList.FocusedRow = Me.rptPatientList.Rows(Me.rptPatientList.Rows.Count - 1)
            Else
                Me.rptPatientList.FocusedRow = Me.rptPatientList.Rows(intIndex)
            End If
            mlngPatienID = Me.rptPatientList.FocusedRow.Record(mCol.病人ID).Value
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptPatientList_SelectionChanged()
    If Me.rptPatientList.FocusedRow Is Nothing Then Exit Sub
    mlngPatienID = Me.rptPatientList.FocusedRow.Record(mCol.病人ID).Value
    mfrmWrite.zlRefreshPatient (mlngPatienID)
End Sub
Private Function AuditingPatient(lngPatientID As Long) As Boolean
    '----------------------------------------
    '功能   按病人为单位进行审核
    '参数   lngPatientID=病人ID
    '--------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New adodb.Recordset
    Dim rs As New adodb.Recordset
    Dim lngKey As Long
    Dim strStart As String, strEnd As String
    Dim intPrivacy As Integer
    Dim blnRollBack As Boolean
    Dim strErrInfo As String
    Dim astrSQL() As String
    Dim intLoop As Integer
    ReDim astrSQL(0)
    
    If InStr(1, mstrPrivs, "审核标本") <= 0 Then
        '没有权限和其他用户登陆时退出
        MsgBox "你没有权限进行审核,请重新登陆具有审核人员进行审核!", vbInformation, gstrSysName
        Exit Function
    End If

            
    '11210 权限“未收费审核”，在审核单个病人时，未生效，
    If InStr(mstrPrivs, "未收费审核") <= 0 Then
        If CheckChargeState(mlngKey, False) = False Then
            MsgBox "单据未收费，不能进行审核！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strStart = GetDateTime(Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(1), 2)
    
     If strStart = "自定义" Then
        strStart = Format(Me.DTPDate.Value, "yyyy-mm-dd 00:00:00")
        strEnd = Format(Me.dtpDateEnd.Value, "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    End If
    
    intPrivacy = zlDatabase.GetPara("报告单是否显示隐私项目", 100, 1208, 0)
    
    strSQL = "select id,检验人 from 检验标本记录 where 病人id = [1] and 核收时间 between [2] and [3] and 医嘱id is not null and 审核人 is null and 微生物标本 is null "
        
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID, CDate(strStart), CDate(strEnd))
    
    
    Do While Not rs.EOF
        lngKey = rs("ID")
        '21137 已归档报告不能审核
        gstrSql = "Select Decode(病案状态, 1, '1-等待审查', 2, '2-拒绝审查', 3, '3-正在审查', 4, '4-审查反馈', 5, '5-审查归档') As 病案状态" & vbNewLine & _
                "From 检验标本记录 A, 病案主页 B ,病案提交记录 C" & vbNewLine & _
                "Where A.病人id = B.病人id And A.主页id = B.主页id And A.病人来源 = 2 And Nvl(B.病案状态, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                " And b.病人id = c.病人Id and B.主页id = C.主页ID "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)
        If rsTmp.EOF = False Then
            MsgBox "病人本次住院的病案已提交审查，不能进行审核！", vbInformation, Me.Caption
'            gcnOracle.RollbackTrans
'            blnRollBack = False
            Exit Function
        End If
                
        '检查住院病人是否出院后还有划价单
        If CheckExesState(lngKey) = False Then
            MsgBox "当前住院病人还有划价单未审核，但已出院或预出院！", vbInformation, Me.Caption
'            gcnOracle.RollbackTrans
'            blnRollBack = False
            Exit Function
        End If
                
        
                
        '检验审核规则判断
        strErrInfo = ""
        If VerifyAuditingRule(lngKey, strErrInfo) = 1 Then
            If Mid(strErrInfo, 1, 2) = "1|" And InStr(mstrPrivs, "强制审核规则") <= 0 Then
                strErrInfo = Mid(strErrInfo, 3)
                MsgBox "<" & strPatienName & ">的检验单审核未通过!" & vbNewLine & strErrInfo
'                gcnOracle.RollbackTrans
'                blnRollBack = False
                Exit Function
            End If
            strErrInfo = Mid(strErrInfo, 3)
            If MsgBox("<" & strPatienName & ">的检验单审核未通过!是否续继?" & vbNewLine & strErrInfo, _
                vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
'                gcnOracle.RollbackTrans
'                blnRollBack = False
                Exit Function
            End If
        End If
                
        
        '签名不成功时退出
        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
        astrSQL(UBound(astrSQL)) = "Signature;" & lngKey & ";" & mstrAuditingMan
        
        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
        astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_报告审核(" & lngKey & ",'" & UserInfo.姓名 & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        
        
        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
        astrSQL(UBound(astrSQL)) = "Zl_检验报告单_Update(" & lngKey & "," & intPrivacy & ",'" & gstrUnitName & "')"             '审核后处理病历报告单
        

        
        
        rs.MoveNext
    Loop
'    gcnOracle.BeginTrans
'    blnRollBack = True
    For intLoop = 1 To UBound(astrSQL)
        If UCase(Mid(astrSQL(intLoop), 1, 3)) = "ZL_" Then
            zlDatabase.ExecuteProcedure astrSQL(intLoop), Me.Caption
        Else
            If Signature(Val(Split(astrSQL(intLoop), ";")(1)), mstrAuditingMan) = False Then
'                gcnOracle.RollbackTrans
'                blnRollBack = False
                Exit Function
            End If
        End If
    Next
'    gcnOracle.CommitTrans
'    If blnAutoPrint Then ReportPrint True                                           '是否完成后直接打印报告
    Exit Function
errH:
    If blnRollBack = True Then
        blnRollBack = False
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function
Private Sub AllReportPrint(lngPatient As Long, blnPrint As Boolean)
    '功能           '按病人打印报告单
    '               lngPatient=病人ID
    '               blnPrint =True打印 False=预览
    
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New adodb.Recordset
    Dim blnCurrMoved As Boolean
    Dim lng医嘱ID As Long, lng发送号 As Long, lng病人ID As Long
    Dim strSQL As String
    Dim strChart(1 To 9) As String
    Dim intLoop As Integer
    Dim lngKey As Long
    Dim str医嘱ID As String                 '医嘱ID，多个医嘱ID使用","分隔。
    Dim str标本ID As String                 '标本ID, 多个标本ID使用","分隔。
    Dim strPrintCode As String              '单据编码
    Dim intItem As Integer
    Dim astrItem() As String
    Dim blnRollBack As Boolean                              '是否回滚
   
    On Error GoTo errH
    
    
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "正在打印请等待...", Me
    
    strStart = GetDateTime(Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(1), 2)
    
    If strStart = "自定义" Then
        strStart = Format(Me.DTPDate, "yyyy-mm-dd 00:00:00")
        strEnd = Format(Me.dtpDateEnd, "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    End If
    
    If strStart = "" Then strStart = GetDateTime("今  天", 1)
    If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    
    strSQL = "Select id,医嘱ID from 检验标本记录 where 病人id = [1] and 核收时间 between [2] and [3] and 医嘱id is not null and 审核人 is null and 微生物标本 is null "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatient, CDate(strStart), CDate(strEnd))
    
    Do While Not rsTmp.EOF
        str医嘱ID = str医嘱ID & "," & rsTmp("医嘱ID")
        str标本ID = str标本ID & "," & rsTmp("ID")
        rsTmp.MoveNext
    Loop
       
    If str医嘱ID <> "" Then str医嘱ID = Mid(str医嘱ID, 2)
    If str标本ID <> "" Then str标本ID = Mid(str标本ID, 2)
    
    lng医嘱ID = Split(str医嘱ID, ",")(0)
    lngKey = Split(str标本ID, ",")(0)
    
    '有多个格式时得到格式
    frmLabMainPrintFormat.ShowMe Me, str医嘱ID, strPrintCode
    
    strSQL = "select /*+ rule */ id from 检验图像结果 where 标本id In(Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str标本ID)
    intLoop = 1
    Do Until rsTmp.EOF
        If intLoop < 9 Then
            strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
            Call LoadImageData(App.path, rsTmp("ID"))
            intLoop = intLoop + 1
        End If
        rsTmp.MoveNext
    Loop
    
    Call ReportOpen(gcnOracle, glngSys, strPrintCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & str医嘱ID, _
                        "病人ID=" & lng病人ID, "标本ID=" & str标本ID, "多个医嘱=" & str医嘱ID, "多个标本=" & str标本ID, _
                        "图形1=" & strChart(1), "图形2=" & strChart(2), "图形3=" & strChart(3), "图形4=" & strChart(4), _
                        "图形5=" & strChart(5), "图形6=" & strChart(6), "图形7=" & strChart(7), "图形8=" & strChart(8), _
                        "图形9=" & strChart(9), IIf(blnPrint, 2, 1))
    
   astrItem = Split(str标本ID, ",")
   gcnOracle.BeginTrans
   blnRollBack = True
   For intLoop = 0 To UBound(astrItem)
        strSQL = "ZL_检验标本记录_标本质控(" & astrItem(intLoop) & ",'',1)"
        zlDatabase.ExecuteProcedure strSQL, gstrSysName
   Next
   gcnOracle.CommitTrans
    Me.MousePointer = 0
    zlCommFun.StopFlash
    
    On Error Resume Next
    '删除图形文件
    For intLoop = 1 To 9
        Kill strChart(intLoop)
    Next
    Exit Sub
errH:
    If blnRollBack = True Then
        gcnOracle.RollbackTrans
    End If
    Me.MousePointer = 0
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMe(objfrm As Object, strPrivs As String, strAuditingMan As String)
    '打开窗口并传入权限
    mstrPrivs = strPrivs
    mstrAuditingMan = strAuditingMan
    Me.Show vbModal, objfrm
End Sub
