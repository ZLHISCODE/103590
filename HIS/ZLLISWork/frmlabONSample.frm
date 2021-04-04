VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmlabONSample 
   Caption         =   "标本保存"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11580
   Icon            =   "frmlabONSample.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   2535
      Left            =   3270
      TabIndex        =   0
      Top             =   1470
      Width           =   3255
      _Version        =   589884
      _ExtentX        =   5741
      _ExtentY        =   4471
      _StockProps     =   0
      BorderStyle     =   1
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   11265
      TabIndex        =   1
      Top             =   570
      Width           =   11295
      Begin VB.OptionButton optSave 
         Caption         =   "待保存"
         Height          =   195
         Index           =   0
         Left            =   8580
         TabIndex        =   4
         Top             =   210
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optSave 
         Caption         =   "已保存"
         Height          =   195
         Index           =   1
         Left            =   9570
         TabIndex        =   3
         Top             =   210
         Width           =   1125
      End
      Begin VB.ComboBox cboMachine 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   150
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker DtpBegin 
         Height          =   285
         Left            =   4260
         TabIndex        =   5
         Top             =   165
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   99745795
         CurrentDate     =   39198
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   285
         Left            =   6420
         TabIndex        =   6
         Top             =   150
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   99745795
         CurrentDate     =   39198
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   6120
         TabIndex        =   9
         Top             =   210
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "审核时间:"
         Height          =   180
         Left            =   3390
         TabIndex        =   8
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "检验仪器:"
         Height          =   180
         Left            =   210
         TabIndex        =   7
         Top             =   210
         Width           =   810
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":68BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":6E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":73F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":798C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":E1EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":14A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":1B2B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":21B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":28376
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   720
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmlabONSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol           '列表
    选择 = 0
    标本号
    标本类型
    病人姓名
    病人来源
    审核时间
    审核人
    保存人
    保存时间
    保存位置
    保存环境
    标本id
End Enum

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Manage_ThingModi               '全选
            Call RptSelect(Me.rptList.Records, True)
            Me.rptList.Populate
        Case conMenu_Manage_ThingDel                '全清
            Call RptSelect(Me.rptList.Records, False)
            Me.rptList.Populate
        Case conMenu_Edit_Import                    '保存
            Call SaveData
        Case conMenu_LIS_Cancel                     '取消保存
            Call SaveData(1)
        Case conMenu_View_Refresh                   '刷新
            Call RefreshData
        Case conMenu_File_Exit                      '退出
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim Column As ReportColumn
    Dim strSQL As String
    Dim rsTmp As New adodb.Recordset
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = zlCommFun.GetPubIcons
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
    Me.cbrthis.ActiveMenuBar.Visible = False
    

    '快键绑定
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_Select
        .Add FCONTROL, Asc("Z"), conMenu_Edit_DeSelect
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F2, conMenu_Edit_Audit
        .Add FCONTROL, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With

    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbrthis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "全选"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "全清")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "存储"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "取消存储")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): cbrControl.BeginGroup = True
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '快键绑定
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_Select
        .Add FCONTROL, Asc("Z"), conMenu_Edit_DeSelect
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F2, conMenu_Edit_Audit
        .Add FCONTROL, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    
    DtpBegin = Now
    DTPEnd = Now
    
    On Error GoTo errH
    
    strSQL = "select id,编码,名称 from 检验仪器 "
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboMachine.Clear
    cboMachine.AddItem "所有仪器"
    cboMachine.ItemData(cboMachine.NewIndex) = 0
    Do Until rsTmp.EOF
        cboMachine.AddItem rsTmp("编码") & "-" & rsTmp("名称")
        cboMachine.ItemData(cboMachine.NewIndex) = rsTmp("ID")
        rsTmp.MoveNext
    Loop
    cboMachine.ListIndex = 0
    
    With Me.rptList.Columns
        
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
        rptList.SetImageList ImgList
        Set Column = .Add(mCol.选择, "选择", 18, False): Column.Icon = 0
        Set Column = .Add(mCol.标本号, "标本号", 80, True)
        Set Column = .Add(mCol.病人姓名, "病人姓名", 65, True)
        Set Column = .Add(mCol.标本类型, "标本类型", 65, True)
        Set Column = .Add(mCol.审核时间, "审核时间", 80, True)
        Set Column = .Add(mCol.审核人, "审核人", 65, True)
        Set Column = .Add(mCol.保存人, "保存人", 65, True)
        Set Column = .Add(mCol.保存时间, "保存时间", 100, True)
        Set Column = .Add(mCol.保存位置, "保存位置", 65, True)
        Set Column = .Add(mCol.保存环境, "保存环境", 65, True)
        Set Column = .Add(mCol.标本id, "标本id", 65, True): Column.Visible = False
        Me.rptList.Populate
    End With
    Call RefreshData
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    With picFilter
        .Top = 460
        .Left = 10
        .Width = Me.ScaleWidth - 25
    End With
    With rptList
        .Top = picFilter.Top + picFilter.Height + 20
        .Left = 10
        .Width = Me.ScaleWidth - 25
        .Height = Me.ScaleHeight - .Top - 20
    End With
End Sub

Private Sub RefreshData()
    '功能           刷新数据
    Dim rsTmp As New adodb.Recordset
    Dim Record As ReportRecord
    Dim intLoop As Integer, lngLoop As Long
    Dim cbrControl As CommandBarControl                 '文本标签
    Dim lngMachineID As Long                            '仪器ID
        
    On Error GoTo errH
    
    If DtpBegin.Value > DTPEnd.Value Then
        MsgBox "开始日期不能大于结束日期！", vbInformation, gstrSysName
        DtpBegin.SetFocus
        Exit Sub
    End If
    
    gstrSql = "select /*+ RULE */ DISTINCT B.相关ID AS ID,A.医嘱id,F.发送号,0 AS 选择," & _
             " Decode(A.仪器id, Null, " & vbCrLf & _
               " to_Char(Trunc(A.标本序号/10000)+1,'0000')|| '-'||to_Char(MOD(A.标本序号,10000),'0000'), to_number(A.标本序号)) As 标本号, " & _
             "A.标本类型," & _
             "TO_CHAR(A.审核时间,'MM-DD HH24:MI') AS 审核时间," & _
             "A.审核人," & _
             "A.检验人," & _
             "TO_CHAR(B.开嘱时间,'MM-DD HH24:MI') AS 申请时间," & _
             "B.开嘱医生 AS 申请人," & _
             "C.名称 AS 申请科室," & _
             "E.名称 AS 执行科室," & _
             "A.id as 标本ID, " & _
             "B.病人id, " & _
             "D.名称 AS 检验仪器,0 As 转出,Decode(A.标本类别,1,'√','') As 急诊, " & _
             "decode(a.审核时间,Null,'否','是') as 是否审核, " & _
             "Decode(a.样本状态, 1, '检验中', 2, '已检验') As 执行状态, " & _
             "Decode(a.是否传送, 1, '', '传送失败') As 传送, a.打印次数,a.微生物标本, " & _
             "a.姓名,a.标本序号,a.仪器ID,a.病人来源,a.婴儿,b.开嘱科室ID,a.报告结果,b.主页ID,保存人,保存时间,保存位置,保存环境  " & _
        "from 检验标本记录 A, 病人医嘱记录 B, 部门表 C, 检验仪器 D,部门表 E,病人医嘱发送 F,病人信息 G " & _
        " WHERE A.医嘱ID = B.相关ID(+) AND B.开嘱科室ID = C.ID(+) AND B.ID=F.医嘱id(+) AND " & _
             "A.仪器ID = D.ID(+) AND B.执行科室id = E.ID AND A.样本状态 IN (1,2) AND a.病人ID = G.病人ID and 审核人 is not null and 销毁人 is null  "
    
    gstrSql = gstrSql & " and 检验时间 between [1] and [2] "
    
    If optSave(0).Value = True Then
        gstrSql = gstrSql & " and 保存人 is null "
        Set cbrControl = cbrthis.FindControl(, conMenu_Edit_Import, True, True)
        cbrControl.Enabled = True
        Set cbrControl = cbrthis.FindControl(, conMenu_LIS_Cancel, True, True)
        cbrControl.Enabled = False
    Else
        gstrSql = gstrSql & " and 保存人 is not null   "
        Set cbrControl = cbrthis.FindControl(, conMenu_Edit_Import, True, True)
        cbrControl.Enabled = False
        Set cbrControl = cbrthis.FindControl(, conMenu_LIS_Cancel, True, True)
        cbrControl.Enabled = True
    End If
    
    If cboMachine.ListIndex >= 0 Then
        If Val(cboMachine.ItemData(cboMachine.ListIndex)) > 0 Then
            gstrSql = gstrSql & " and 仪器ID = [3] "
            lngMachineID = cboMachine.ItemData(cboMachine.ListIndex)
        End If
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(Format(Me.DtpBegin, "yyyy-mm-dd 00:00:00")), _
                                         CDate(Format(Me.DTPEnd, "yyyy-mm-dd 23:23:59")), lngMachineID)
    
    
    
    With Me.rptList
        .Records.DeleteAll
        .Populate
        Do Until rsTmp.EOF
            Set Record = .Records.Add
            .Populate
            For intLoop = 0 To .Columns.Count
                Record.AddItem ""
            Next
            Record.Item(mCol.选择).HasCheckbox = True
            Record.Item(mCol.标本id).Value = Nvl(rsTmp("标本ID"))
            Record.Item(mCol.标本号).Value = Val(Nvl(rsTmp("标本序号")))
            Record.Item(mCol.标本号).Caption = Trim(Nvl(rsTmp("标本号")))
            Record.Item(mCol.标本类型).Value = Nvl(rsTmp("标本类型"))
            Record.Item(mCol.病人姓名).Value = Nvl(rsTmp("姓名"))
            Record.Item(mCol.审核人).Value = Nvl(rsTmp("审核人"))
            Record.Item(mCol.审核时间).Value = Nvl(rsTmp("审核时间"))
            Record.Item(mCol.保存人).Value = Nvl(rsTmp("保存人"))
            Record.Item(mCol.保存时间).Value = Nvl(rsTmp("保存时间"))
            Record.Item(mCol.保存位置).Value = Nvl(rsTmp("保存位置"))
            Record.Item(mCol.保存环境).Value = Nvl(rsTmp("保存环境"))
            Record.Item(mCol.标本id).Value = Nvl(rsTmp("标本id"))
            rsTmp.MoveNext
        Loop
        .Populate
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub RptSelect(Records As ReportRecords, blTrue As Boolean)
    '功能                           选择或取消选择
    '参数                           Records = 列表对象
    '                               blTrue  True = 选择 False = 取消选择
    Dim intLoop As Integer
    
    For intLoop = 0 To Records.Count - 1
        Records(intLoop).Item(mCol.选择).Checked = blTrue
    Next
End Sub

Private Sub optSave_Click(Index As Integer)
    Call RefreshData
End Sub

Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, x As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptList
        Set hitColumn = .HitTest(x, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "选择" And .HitTest(x, Y).ht = xtpHitTestHeader Then
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mCol.选择).Checked
                For Each Record In .Records
                    Record.Item(mCol.选择).Checked = blSelect
                Next
            End If
        End If
        .Populate
    End With
End Sub

Private Sub SaveData(Optional intType As Integer)
    '功能           保存
    '参数           intType 0=标本保存 1=标本取消保存
    Dim strVal As String
    Dim astrVal() As String
    Dim strIDs As String
    Dim strSQL As String
    Dim intLoop As Integer
    
    On Error GoTo errH
    
    If CheckSel = False Then
        MsgBox "你一个标本都没有选择不能保存!", vbInformation, "保存标本"
        Exit Sub
    End If
    If intType = 0 Then
        '保存
        strVal = frmlabONSampleEdit.ShowMe(Me)
        If strVal = "" Then Exit Sub
        
        '开始组织数据
        astrVal = Split(strVal, "|")
        With Me.rptList
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mCol.选择).Checked = True Then
                    strIDs = strIDs & "," & .Records(intLoop).Item(mCol.标本id).Value
                End If
            Next
        End With
        
        If strIDs <> "" Then
            strIDs = Mid(strIDs, 2)
            If strIDs <> "" Then
                '保存
                strSQL = "ZL_检验标本保存_edit(0,'" & strIDs & "','" & astrVal(0) & "','" & astrVal(1) & "','" & astrVal(2) & "')"
                zldatabase.ExecuteProcedure strSQL, "保存标本"
            End If
        End If
    Else
        '取消保存
        With Me.rptList
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mCol.选择).Checked = True Then
                    strIDs = strIDs & "," & .Records(intLoop).Item(mCol.标本id).Value
                End If
            Next
        End With
        If strIDs <> "" Then
            strIDs = Mid(strIDs, 2)
            If strIDs <> "" Then
                '保存
                strSQL = "ZL_检验标本保存_edit(1,'" & strIDs & "','','','')"
                zldatabase.ExecuteProcedure strSQL, "保存标本"
            End If
        End If
    End If
    Call RefreshData
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function CheckSel() As Boolean
    '功能           检查当前列表是否有选择的记录
    Dim intLoop As Integer
    With Me.rptList
        For intLoop = 0 To .Records.Count - 1
            If .Records(intLoop).Item(mCol.选择).Checked = True Then
                CheckSel = True
                Exit Function
            End If
        Next
    End With
End Function

